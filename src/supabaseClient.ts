import { createClient } from '@supabase/supabase-js';

// ⚠️ 請將下方的網址與 anon-key 替換為您在 Supabase 後台 Project Settings > API 複製的真實內容
const supabaseUrl = 'https://usjanvounmuatqvzwxxi.supabase.co';
const supabaseAnonKey =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVzamFudm91bm11YXRxdnp3eHhpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAyNzk3MTksImV4cCI6MjA5NTg1NTcxOX0.LahbCbCdx9G8ZMvinzJqfYUUeweVt_C98tgHCun6GjM';

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
