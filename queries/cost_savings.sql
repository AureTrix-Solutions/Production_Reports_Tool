SELECT
  Max(mom_fgc) AS fgc,
  Max(mom_tec) AS tec,
  Max(mom_niin) AS niin,
  mom_mcn AS mcn,
  Max(wc_cd) AS wc,
  Max(mom_status_date) AS status_date,
  Max(mom_repair_cost) AS net_cost,
  Sum(repair_part_cost) AS frc_repair_cost,
  (
    Max(mom_repair_cost) - Sum(repair_part_cost)
  ) AS net_cost_savings
FROM
  Cost_savings_work
GROUP BY
  mom_mcn
HAVING
  Year(
    Max(mom_status_date)
  ) >= (
    Year(
      Date()
    ) - 2
  );
