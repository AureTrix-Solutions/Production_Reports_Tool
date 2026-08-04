SELECT
  c.fgc,
  c.niin,
  hof_niin.hof_niin,
  c.wc_cd,
  c.mcn,
  c.jcn,
  Left(c.jcn, 3) AS org_cd,
  c.cur_mgmt_cd,
  c.e_part_no,
  c.e_serno,
  c.sys_rsn,
  c.current_status,
  c.Maf_age
FROM
  combined AS c
  LEFT JOIN hof_niin ON c.fgc = hof_niin.fgc
WHERE
  c.current_status = "iw"
  AND (
    c.wc_cd LIKE "6*"
    OR c.wc_cd LIKE "5*"
  )
ORDER BY
  c.wc_cd;
