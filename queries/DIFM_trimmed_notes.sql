SELECT
  combined.fgc,
  combined.niin,
  combined.wc_cd,
  combined.mcn,
  combined.jcn,
  Left([jcn], 3) AS org_cd,
  combined.cur_mgmt_cd,
  combined.e_part_no,
  combined.e_serno,
  combined.sys_rsn,
  combined.current_status,
  combined.Maf_age
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
    AND (
      (combined.ty_maf_cd) = "d"
    )
    AND (combined.se_part_no) IS NULL
    AND (combined.cur_ti_ddsn) IS NOT NULL
    AND (combined.wc_cd) NOT LIKE ("*y")
    AND (combined.wc_cd) NOT LIKE ("*z")
  );
