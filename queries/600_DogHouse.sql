SELECT
  600 AS Div,
  T1.fgc,
  T1.tec,
  T1.wc_cd,
  T1.mcn,
  T1.jcn,
  T1.niin,
  t1.hof_niin AS hof,
  t1.e_part_no,
  t1.e_serno,
  t1.se_part_no,
  t1.buno_serno,
  t1.sys_rsn,
  T1.initiated_date,
  T1.current_status,
  t1.current_status_date,
  t1.ty_maf_cd,
  T1.Maf_age
FROM
  600_D AS T1
  LEFT JOIN 600_d AS T2 ON (
    T1.Maf_age < T2.Maf_age
    OR (
      T1.Maf_age = T2.Maf_age
      AND T1.mcn < T2.mcn
    )
  )
  AND (T1.wc_cd = T2.wc_cd)
WHERE
  T1.wc_cd LIKE "6*"
  AND T1.current_status LIKE "m*"
GROUP BY
  T1.wc_cd,
  T1.Maf_age,
  T1.mcn,
  T1.jcn,
  T1.fgc,
  T1.tec,
  T1.niin,
  t1.hof_niin,
  T1.initiated_date,
  T1.current_status,
  T1.ty_maf_cd,
  t1.sys_rsn,
  t1.e_part_no,
  t1.e_serno,
  t1.buno_serno,
  t1.current_status_date,
  t1.se_part_no
HAVING
  Count(T2.mcn) < 5
ORDER BY
  T1.wc_cd,
  T1.Maf_age DESC;
