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
  combined.proj_cd
FROM
  combined
  LEFT JOIN hof_niin ON combined.fgc = hof_niin.fgc
WHERE
  (
    (
      (combined.cur_mgmt_cd) <> "SO"
      OR (cur_mgmt_cd) IS NULL
    )
    AND (
      (combined.current_status) <> "jc"
      AND (combined.current_status) <> "rj"
      AND (combined.current_status) <> "cc"
    )
    AND (
      Left([jcn], 3) <> "P9H"
    )
    AND (cntl_mcn) IS NULL
    AND (
      (wc_cd) IS NOT NULL
      AND (wc_cd) NOT IN ("64B", "64G", "64E", "64T", "64Q")
      AND (wc_cd) NOT LIKE ("*Z")
      AND (wc_cd) NOT LIKE ("9*")
      AND (wc_cd) NOT LIKE ("8*")
      AND (wc_cd) NOT LIKE ("7*")
      AND (wc_cd) NOT LIKE ("x*")
      AND (wc_cd) NOT LIKE ("0*")
    )
    AND (
      (tec) <> "BAED"
    )
    AND (
      (proj_cd) IS NULL
      OR (proj_cd) NOT IN ("z6z", "730")
    )
  );
