TRANSFORM
  Count([500_600_BS].tec) AS CountOftec
SELECT
  [500_600_BS].fgc,
  MAX([500_600_BS].hof_niin) AS hof_niin,
  [500_600_BS].wc_cd,
  Max([500_600_BS].tec) AS tec,
  Count([500_600_BS].tec) AS difm,
  Max([500_600_BS].exrp_qty) AS exrep,
  Max([500_600_BS].supply_percent) AS supply,
  Max([500_600_BS].buffer_status) AS buffer
FROM
  500_600_BS
WHERE
  (
    (
      ([500_600_BS].fgc) Is Not Null
    )
  )
GROUP BY
  [500_600_BS].fgc,
  [500_600_BS].wc_cd
PIVOT
  [500_600_BS].current_status;
