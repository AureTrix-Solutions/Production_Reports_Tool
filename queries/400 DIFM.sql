SELECT
  combined.fgc,
  combined.tec,
  combined.niin,
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
  Buffer_status.buffer_status,
  Buffer_status.rfi_qty,
  Buffer_status.difm_qty,
  Buffer_status.exrp_qty,
  buffer_status.supply_percent
FROM
  combined
  LEFT JOIN Buffer_status ON combined.fgc = Buffer_status.fgc
WHERE
  (
    (
      (combined.wc_cd) LIKE "4**"
    )
    AND (
      (combined.current_status) <> "jc"
      AND (combined.current_status) <> "rj"
      AND (combined.current_status) <> "cc"
    )
  );
