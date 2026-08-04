UPDATE
  combined
SET
  combined.cost_key = [wc_cd] & [jcn]
WHERE
  (
    (
      (combined.cost_key) Is Null
    )
  );
