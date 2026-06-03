/**
 * Low-level OAuth 2.1 helpers: secure random token generation, PKCE (RFC 7636)
 * generation and verification, and redirect-URI validation.
 *
 * All randomness uses Node's crypto.randomBytes; all comparisons of secrets use
 * timing-safe equality where a secret is involved.
 */
import { randomBytes, createHash, timingSafeEqual } from "crypto";

/** Base64url-encode a buffer (no padding). */
export function base64url(buf: Buffer): string {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/** Generate a high-entropy opaque token (default 32 bytes → 43 base64url chars). */
export function generateToken(bytes = 32): string {
  return base64url(randomBytes(bytes));
}

/** SHA-256 of a string, returned as base64url — used for PKCE S256 and as a
 *  storage key for opaque tokens (so we never persist the raw secret). */
export function sha256(value: string): string {
  return base64url(createHash("sha256").update(value).digest());
}

/** Generate a PKCE verifier/challenge pair (S256) for our leg to Entra. */
export function generatePkce(): { verifier: string; challenge: string } {
  const verifier = generateToken(32);
  const challenge = sha256(verifier);
  return { verifier, challenge };
}

/**
 * Verify a PKCE code_verifier against a stored S256 code_challenge.
 * Timing-safe to avoid leaking challenge bytes.
 */
export function verifyPkceS256(verifier: string, challenge: string): boolean {
  const computed = sha256(verifier);
  const a = Buffer.from(computed);
  const b = Buffer.from(challenge);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

/**
 * Validate a client-supplied redirect URI against the set registered for the
 * client. Exact string match only — no prefix/substring matching, which is a
 * classic open-redirect footgun.
 */
export function isRegisteredRedirectUri(uri: string, registered: string[]): boolean {
  return registered.includes(uri);
}

/**
 * Sanity-check a redirect URI at registration time. We allow https everywhere,
 * plus http only for loopback (native/desktop clients), per OAuth 2.1 BCP.
 */
export function isAcceptableRedirectUri(uri: string): boolean {
  let u: URL;
  try {
    u = new URL(uri);
  } catch {
    return false;
  }
  if (u.protocol === "https:") return true;
  if (u.protocol === "http:" && (u.hostname === "127.0.0.1" || u.hostname === "localhost")) {
    return true;
  }
  return false;
}

/** Build a redirect back to the client, appending code+state to its redirect URI. */
export function buildClientRedirect(redirectUri: string, params: Record<string, string>): string {
  const u = new URL(redirectUri);
  for (const [k, v] of Object.entries(params)) {
    u.searchParams.set(k, v);
  }
  return u.toString();
}
