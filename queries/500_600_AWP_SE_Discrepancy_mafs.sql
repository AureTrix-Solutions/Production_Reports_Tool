SELECT
  combined.tec,
  combined.wc_cd,
  combined.mcn,
  combined.jcn,
  combined.e_part_no,
  combined.e_serno,
  combined.se_part_no,
  combined.buno_serno,
  combined.sys_rsn,
  combined.initiated_date,
  combined.current_status,
  combined.current_status_date,
  combined.Maf_age,
  combined.ty_maf_cd,
  combined.maint_lv_cd,
  combined.wrk_pri_cd
FROM
  combined
  LEFT JOIN Buffer_status ON combined.fgc = Buffer_status.fgc
WHERE
  (
    (
      combined.wc_cd LIKE ("5*")
      AND (
        (combined.current_status) <> "jc"
        AND (combined.current_status) <> "rj"
        AND (combined.current_status) <> "cc"
      )
      OR (
        combined.wc_cd LIKE ("6*")
        AND (
          (combined.current_status) <> "jc"
          AND (combined.current_status) <> "rj"
          AND (combined.current_status) <> "cc"
        )
      )
    )
    AND (combined.se_part_no) IS NOT NULL
    AND (combined.cur_ti_ddsn) IS NULL
    AND (combined.ty_maf_cd) = "d"
    AND (combined.wc_cd) NOT LIKE ("*z")
    AND (combined.wc_cd) NOT LIKE ("*y")
    AND (combined.current_status) LIKE ("w*")
  )
ORDER BY
  wc_cd,
  tec;
