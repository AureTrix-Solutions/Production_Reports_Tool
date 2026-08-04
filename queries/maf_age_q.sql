UPDATE
  combined
SET
  combined.Maf_age = IIf(
    [current_status] Like "JC",
    DateDiff(
      "d", [initiated_date], [current_status_date]
    ),
    DateDiff(
      "d",
      [initiated_date],
      Now()
    )
  );
