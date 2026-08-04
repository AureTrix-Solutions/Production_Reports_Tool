SELECT
  combined.fgc,
  combined.niin,
  combined.wc_cd,
  combined.mcn,
  Left([jcn], 3) AS org_cd,
  combined.jcn,
  combined.e_part_no,
  combined.e_serno,
  combined.se_part_no,
  combined.buno_serno,
  combined.sys_rsn,
  combined.unit_price,
  combined.net_price,
  combined.maint_trans_cd,
  combined.act_take_cd,
  combined.A1_status,
  combined.A1_date,
  combined.current_status,
  combined.current_status_date,
  DateDiff(
    "d", [A1_date], [current_status_date]
  ) AS days_on_backlog,
  combined.ty_maf_cd,
  combined.maint_lv_cd,
  combined.wrk_pri_cd
FROM
  combined;
