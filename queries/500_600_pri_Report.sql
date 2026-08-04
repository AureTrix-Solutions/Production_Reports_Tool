SELECT
  *
FROM
  500_600_pri_report_crosstab
WHERE
  (
    (
      (
        [500_600_pri_report_crosstab].[wc_cd]
      ) Not Like ("*Y")
      And (
        [500_600_pri_report_crosstab].[wc_cd]
      ) Not Like ("*z")
    )
  )
ORDER BY
  [500_600_pri_report_crosstab].wc_cd,
  [500_600_pri_report_crosstab].supply,
  [500_600_pri_report_crosstab].exrep DESC,
  [500_600_pri_report_crosstab].difm DESC;
