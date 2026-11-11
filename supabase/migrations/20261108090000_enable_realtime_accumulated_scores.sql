/*
  # Enable realtime for accumulated_scores
  Lets the app push score changes (daily save, exchange, import) to all open screens instantly.
*/
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables
    WHERE pubname = 'supabase_realtime' AND schemaname = 'public' AND tablename = 'accumulated_scores'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.accumulated_scores;
  END IF;
END $$;
