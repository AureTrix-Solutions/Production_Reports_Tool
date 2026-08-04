SELECT
  combined.fgc,
  combined.tec,
  combined.niin,
  hof_niin.hof_niin,
  combined.wc_cd,
  combined.mcn,
  combined.cur_ti_ddsn,
  Left([jcn], 3) AS org_cd,
  combined.cur_mgmt_cd,
  combined.jcn,
  combined.e_part_no,
  combined.e_serno,
  combined.se_part_no,
  combined.buno_serno,
  combined.sys_rsn,
  combined.unit_price,
  combined.net_price,
  combined.act_take_cd,
  combined.maint_trans_cd,
  combined.initiated_date,
  combined.current_status,
  combined.current_status_date,
  combined.Maf_age,
  combined.ty_maf_cd,
  combined.maint_lv_cd,
  combined.wrk_pri_cd,
  combined.proj_cd,
  [fgc_bsr_work].buffer_status,
  [fgc_bsr_work].design_bsz,
  [fgc_bsr_work].faq,
  [fgc_bsr_work].acbal,
  [fgc_bsr_work].rfi_qty,
  [fgc_bsr_work].difm_qty,
  [fgc_bsr_work].exrp_qty,
  [fgc_bsr_work].due_qty,
  [fgc_bsr_work].owed_qty,
  [fgc_bsr_work].supply_percent,
  [fgc_bsr_work].pkup_qty
FROM
  (
    combined
    LEFT JOIN fgc_bsr_work ON combined.fgc = [fgc_bsr_work].fgc
  )
  LEFT JOIN hof_niin ON combined.fgc = hof_niin.fgc
WHERE
  combined.current_status NOT IN ("jc", "rj", "cc")
  AND combined.ty_maf_cd = "d"
  AND [fgc_bsr_work].buffer_status = "red"
  AND combined.cur_ti_ddsn IS NOT NULL
  AND combined.wc_cd NOT LIKE "*y"
  AND combined.wc_cd NOT LIKE "*z"
  AND combined.wc_cd <> "05a";
