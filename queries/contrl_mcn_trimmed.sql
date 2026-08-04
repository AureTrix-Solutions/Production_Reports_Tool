SELECT
  combined.wc_cd,
  combined.mcn,
  combined.cntl_mcn,
  combined.sys_rsn,
  combined.current_status,
  combined.Maf_age
FROM
  combined
WHERE
  (
    (
      (combined.current_status) <> "jc"
      AND (combined.current_status) <> "rj"
      AND (combined.current_status) <> "cc"
    )
    AND (
      (combined.[cntl_mcn]) IS NOT NULL
    )
    AND (
      (combined.[ty_maf_cd]) LIKE "A*"
      OR ty_maf_cd LIKE "P*"
      OR ty_maf_cd LIKE "s*"
    )
  );
