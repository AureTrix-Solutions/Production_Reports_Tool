SELECT
  Sum(acct_qty) AS acbal_total,
  Sum(depl_pkup_qty) AS pkup_qty_total,
  Sum(difm_qty) AS difm_qty_total,
  Sum(exrp_qty) AS exrp_qty_total,
  fgc,
  Sum(owe_qty) AS owed_qty_total,
  Sum(rfi_qty) AS rfi_qty_total,
  Sum(subcus_qty) AS subcustody_qty,
  Sum(wpurp_faq) AS faq_total,
  Sum(due_qty) AS due_qty_total
FROM
  Inventory
GROUP BY
  fgc
HAVING
  fgc NOT LIKE "";
