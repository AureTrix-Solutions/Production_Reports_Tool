SELECT
  *
FROM
  500_600_bs_crosstab
WHERE
  (
    (
      ([500_600_bs_crosstab].[wc_cd]) Not Like ("*Y")
      And ([500_600_bs_crosstab].[wc_cd]) Not Like ("*z")
    )
  )
ORDER BY
  [500_600_bs_crosstab].wc_cd,
  [500_600_bs_crosstab].supply_percnt,
  [500_600_bs_crosstab].exrep_qty DESC,
  [500_600_bs_crosstab].difm DESC;
