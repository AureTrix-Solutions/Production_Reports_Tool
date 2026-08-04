SELECT
  combined.fgc,
  combined.tec,
  combined.niin,
  combined.wc_cd,
  combined.mcn,
  combined.cntl_mcn,
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
  combined.proj_cd
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
      or ty_maf_cd LIKE "P*"
      or ty_maf_cd LIKE "s*"
    )
  );
