import { createClient } from '@supabase/supabase-js'

const SUPABASE_URL = 'https://kbfttuwblanvkiocgktf.supabase.co'
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtiZnR0dXdibGFudmtpb2Nna3RmIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODA3NTYxMDEsImV4cCI6MjA5NjMzMjEwMX0.-lCHjSm674e0ePV8fMZmwDrX_XFulxjNy7vA0VC_fqE'

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY)