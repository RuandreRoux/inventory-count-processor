import { createClient, type SupabaseClient } from '@supabase/supabase-js';

export const SUPABASE_ENABLED = Boolean(
  process.env.SUPABASE_URL && process.env.SUPABASE_SERVICE_ROLE_KEY
);

let _client: SupabaseClient | null = null;

export function getSupabase(): SupabaseClient {
  if (!_client) {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
    if (!url || !key) {
      throw new Error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY env vars are required');
    }
    _client = createClient(url, key, { auth: { persistSession: false } });
  }
  return _client;
}
