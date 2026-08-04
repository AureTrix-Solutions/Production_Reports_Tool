SELECT
  fgc,
  MAX(hof_niin) AS hof,
  MAX(tec) AS tec_code,
  Count(fgc) AS AK0_AK7_difm,
  MAX(exrp_qty) AS exreps,
  MAX(owed_qty) AS owed,
  MAX(rfi_qty) AS rfi,
  MAX(design_bsz) AS buffer_design,
  MAX(faq) AS faq_qty,
  MAX(acbal) AS acbal_qty,
  MAX(due_qty) AS due_in,
  MAX(pkup_qty) AS pkup,
  MAX(proj_cd) AS proj
FROM
  [MMCO BS]
WHERE
  proj_cd IN ("ak0", "ak7")
GROUP BY
  fgc;
