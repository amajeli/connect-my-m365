import { test } from "node:test";
import assert from "node:assert/strict";
import { isAllowedRedirectUri } from "../oauth";

test("redirect URIs are limited to Claude hosts and loopback", () => {
  assert.ok(isAllowedRedirectUri("https://claude.ai/api/mcp/auth_callback"));
  assert.ok(isAllowedRedirectUri("https://claude.com/api/mcp/auth_callback"));
  assert.ok(isAllowedRedirectUri("http://localhost:6274/oauth/callback"));
  assert.ok(isAllowedRedirectUri("http://127.0.0.1:33418/callback"));

  assert.ok(!isAllowedRedirectUri("https://evil.example/callback"));
  assert.ok(!isAllowedRedirectUri("https://claude.ai.evil.example/cb"));
  assert.ok(!isAllowedRedirectUri("http://claude.ai/api/mcp/auth_callback"));
  assert.ok(!isAllowedRedirectUri("https://localhost/cb"));
  assert.ok(!isAllowedRedirectUri("not a url"));
});

test("extra hosts can be allowed via OAUTH_ALLOWED_REDIRECT_HOSTS", () => {
  process.env.OAUTH_ALLOWED_REDIRECT_HOSTS = "app.example.com";
  try {
    assert.ok(isAllowedRedirectUri("https://app.example.com/cb"));
  } finally {
    delete process.env.OAUTH_ALLOWED_REDIRECT_HOSTS;
  }
});
