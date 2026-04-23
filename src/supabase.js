import { createClient } from '@supabase/supabase-js';

const SUPABASE_URL = process.env.REACT_APP_SUPABASE_URL || 'https://onhhighwzmzwixvtwmpg.supabase.co/';
const SUPABASE_KEY = process.env.REACT_APP_SUPABASE_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9uaGhpZ2h3em16d2l4dnR3bXBnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY3ODk2MjAsImV4cCI6MjA5MjM2NTYyMH0.Fk_UCM97dPw1CrUGcKbPnHQMbZnyU-hks1Sinw-ZRv0';

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);