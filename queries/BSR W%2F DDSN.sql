SELECT
  fgc,
  MAX(hof_niin) AS hof,
  MAX(tec) AS Tec_Code,
  Count(fgc) AS All_BSR_difm,
  MAX(exrp_qty) AS exreps,
  MAX(owed_qty) AS owed,
  MAX(rfi_qty) AS rfi,
  MAX(design_bsz) AS buffer_design,
  MAX(faq) AS faq_qty,
  MAX(acbal) AS acbal_qty,
  MAX(due_qty) AS due_in,
  MAX(pkup_qty) AS pkup
FROM
  [MMCO BS]
GROUP BY
  fgc;
