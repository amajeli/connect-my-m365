import pg from 'pg';
const { Pool } = pg;

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

await pool.query(`
  CREATE TABLE IF NOT EXISTS users (
    id VARCHAR PRIMARY KEY DEFAULT gen_random_uuid(),
    email TEXT NOT NULL UNIQUE,
    display_name TEXT NOT NULL,
    microsoft_id TEXT NOT NULL UNIQUE,
    mcp_api_key TEXT NOT NULL DEFAULT gen_random_uuid(),
    stripe_customer_id TEXT,
    stripe_subscription_id TEXT,
    subscription_status TEXT DEFAULT 'none',
    created_at TIMESTAMP NOT NULL DEFAULT now()
  )
`);

await pool.query(`
  CREATE TABLE IF NOT EXISTS connected_accounts (
    id VARCHAR PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id VARCHAR NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    email TEXT NOT NULL,
    display_name TEXT NOT NULL,
    account_type TEXT NOT NULL,
    access_token TEXT NOT NULL,
    refresh_token TEXT NOT NULL,
    token_expiry TIMESTAMP NOT NULL,
    is_active BOOLEAN NOT NULL DEFAULT true,
    connected_at TIMESTAMP NOT NULL DEFAULT now()
  )
`);

await pool.query(`
  CREATE TABLE IF NOT EXISTS oauth_clients (
    id VARCHAR PRIMARY KEY DEFAULT gen_random_uuid(),
    client_id TEXT NOT NULL UNIQUE,
    client_secret TEXT NOT NULL,
    client_name TEXT,
    redirect_uris TEXT[] NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT now()
  )
`);

await pool.query(`
  CREATE TABLE IF NOT EXISTS oauth_codes (
    id VARCHAR PRIMARY KEY DEFAULT gen_random_uuid(),
    code TEXT NOT NULL UNIQUE,
    client_id TEXT NOT NULL,
    user_id VARCHAR NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    redirect_uri TEXT NOT NULL,
    scope TEXT,
    code_challenge TEXT,
    code_challenge_method TEXT,
    expires_at TIMESTAMP NOT NULL,
    used BOOLEAN NOT NULL DEFAULT false,
    created_at TIMESTAMP NOT NULL DEFAULT now()
  )
`);

await pool.query(`
  CREATE TABLE IF NOT EXISTS oauth_tokens (
    id VARCHAR PRIMARY KEY DEFAULT gen_random_uuid(),
    access_token TEXT NOT NULL UNIQUE,
    refresh_token TEXT UNIQUE,
    client_id TEXT NOT NULL,
    user_id VARCHAR NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    scope TEXT,
    expires_at TIMESTAMP NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT now()
  )
`);

console.log('All tables created successfully');
await pool.end();
