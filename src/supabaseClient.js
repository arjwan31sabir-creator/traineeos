import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://kbfttuwblanvkiocgktf.supabase.co'
const SUPABASE_ANON_KEY = 'sb_publishable_vQIUaiCdIGeOx9yXQHUScQ_N7itSU9Q'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)