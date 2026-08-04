UPDATE
  combined
SET
  combined.A1_date = Left(
    [a1_date_text],
    InStr([a1_date_text], " ") - 1
  ),
  combined.current_status_date = Left(
    [current_status_date_text],
    InStr([current_status_date_text], " ") - 1
  ),
  combined.initiated_date = Left(
    [initiated_date_text],
    InStr([initiated_date_text], " ") - 1
  );
