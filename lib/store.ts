/**
 * Persistence layer for the OAuth authorization server, backed by Upstash Redis
 * (provisioned via the Vercel Marketplace Redis integration). Everything is
 * short-lived except client registrations and user sessions, and Entra tokens
 * are encrypted at rest with AES-256-GCM.
 *
 * Key namespaces:
 *   client:{clientId}        registered DCR clients
 *   authreq:{state}          in-flight authorize requests (10 min)
 *   authcode:{code}          one-time authorization codes (60 s)
 *   session:{sessionId}      user session w/ encrypted Entra tokens (90 d)
 *   atoken:{sha256(bearer)}  access-token -> session map (1 h)
 *   rtoken:{sha256(refresh)} refresh-token -> session map (90 d)
 */
import { Redis } from "@upstash/redis";
import { createCipheriv, createDecipheriv, randomBytes } from "crypto";
import { env } from "./env";

// The Vercel Redis integration injects KV_REST_API_* (and Upstash injects
// UPSTASH_REDIS_REST_*). Accept either naming.
const redis = new Redis({
  url: process.env.KV_REST_API_URL ?? process.env.UPSTASH_REDIS_REST_URL ?? "",
  token: process.env.KV_REST_API_TOKEN ?? process.env.UPSTASH_REDIS_REST_TOKEN ?? "",
});

const ENC_KEY = Buffer.from(env.WCIO_MCP_TOKEN_ENCRYPTION_KEY, "base64");

const TTL = {
  client: 90 * 24 * 60 * 60,
  authReq: 10 * 60,
  authCode: 60,
  session: 90 * 24 * 60 * 60,
  accessToken: 60 * 60,
  refreshToken: 90 * 24 * 60 * 60,
} as const;

// ---------------------------------------------------------------------------
// AES-256-GCM encryption for tokens at rest
// ---------------------------------------------------------------------------

/** Encrypt a UTF-8 string → base64(iv | authTag | ciphertext). */
export function encrypt(plaintext: string): string {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", ENC_KEY, iv);
  const ct = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();
  return Buffer.concat([iv, tag, ct]).toString("base64");
}

/** Decrypt base64(iv | authTag | ciphertext) → UTF-8 string. */
export function decrypt(payload: string): string {
  const raw = Buffer.from(payload, "base64");
  const iv = raw.subarray(0, 12);
  const tag = raw.subarray(12, 28);
  const ct = raw.subarray(28);
  const decipher = createDecipheriv("aes-256-gcm", ENC_KEY, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString("utf8");
}

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface ClientRecord {
  clientId: string;
  redirectUris: string[];
  clientName?: string;
  createdAt: number;
}

export interface AuthRequest {
  clientId: string;
  clientRedirectUri: string;
  clientState?: string;
  clientCodeChallenge: string;
  clientCodeChallengeMethod: string;
  /** Our own PKCE verifier for the leg to Entra. */
  entraVerifier: string;
  createdAt: number;
}

export interface AuthCode {
  clientId: string;
  redirectUri: string;
  clientCodeChallenge: string;
  clientCodeChallengeMethod: string;
  sessionId: string;
}

export interface Session {
  sessionId: string;
  userOid?: string;
  userIdentifier: string;
  /** Encrypted Graph access token. */
  graphAccessTokenEnc: string;
  /** Encrypted Entra refresh token. */
  graphRefreshTokenEnc: string;
  /** Epoch millis when the Graph access token expires. */
  graphExpiresAt: number;
  scopes: string[];
  createdAt: number;
}

export interface TokenRef {
  sessionId: string;
  clientId: string;
  expiresAt: number;
}

// ---------------------------------------------------------------------------
// Clients (Dynamic Client Registration)
// ---------------------------------------------------------------------------

export async function putClient(rec: ClientRecord): Promise<void> {
  await redis.set(`client:${rec.clientId}`, rec, { ex: TTL.client });
}

export async function getClient(clientId: string): Promise<ClientRecord | null> {
  return (await redis.get<ClientRecord>(`client:${clientId}`)) ?? null;
}

// ---------------------------------------------------------------------------
// In-flight authorize requests (keyed by our state)
// ---------------------------------------------------------------------------

export async function putAuthRequest(state: string, req: AuthRequest): Promise<void> {
  await redis.set(`authreq:${state}`, req, { ex: TTL.authReq });
}

export async function takeAuthRequest(state: string): Promise<AuthRequest | null> {
  const key = `authreq:${state}`;
  const req = await redis.get<AuthRequest>(key);
  if (req) await redis.del(key);
  return req ?? null;
}

// ---------------------------------------------------------------------------
// One-time authorization codes (keyed by the code we issue to the client)
// ---------------------------------------------------------------------------

export async function putAuthCode(code: string, rec: AuthCode): Promise<void> {
  await redis.set(`authcode:${code}`, rec, { ex: TTL.authCode });
}

export async function takeAuthCode(code: string): Promise<AuthCode | null> {
  const key = `authcode:${code}`;
  const rec = await redis.get<AuthCode>(key);
  if (rec) await redis.del(key);
  return rec ?? null;
}

// ---------------------------------------------------------------------------
// Sessions
// ---------------------------------------------------------------------------

export async function putSession(s: Session): Promise<void> {
  await redis.set(`session:${s.sessionId}`, s, { ex: TTL.session });
}

export async function getSession(sessionId: string): Promise<Session | null> {
  return (await redis.get<Session>(`session:${sessionId}`)) ?? null;
}

// ---------------------------------------------------------------------------
// Access / refresh token reference maps (keyed by sha256 of the raw token)
// ---------------------------------------------------------------------------

export async function putAccessToken(tokenHash: string, ref: TokenRef): Promise<void> {
  await redis.set(`atoken:${tokenHash}`, ref, { ex: TTL.accessToken });
}

export async function getAccessTokenRef(tokenHash: string): Promise<TokenRef | null> {
  return (await redis.get<TokenRef>(`atoken:${tokenHash}`)) ?? null;
}

export async function putRefreshToken(tokenHash: string, ref: TokenRef): Promise<void> {
  await redis.set(`rtoken:${tokenHash}`, ref, { ex: TTL.refreshToken });
}

export async function takeRefreshToken(tokenHash: string): Promise<TokenRef | null> {
  // Refresh tokens rotate on use, so consume the old one.
  const key = `rtoken:${tokenHash}`;
  const ref = await redis.get<TokenRef>(key);
  if (ref) await redis.del(key);
  return ref ?? null;
}

export const TOKEN_TTL = TTL;
