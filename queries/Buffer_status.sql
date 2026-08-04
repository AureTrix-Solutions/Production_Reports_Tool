SELECT
  fgc,
  (acbal_total - pkup_qty_total) AS design_bsz,
  faq_total AS faq,
  acbal_total AS acbal,
  due_qty_total AS due_qty,
  (rfi_qty_total) AS rfi_qty,
  difm_qty_total AS difm_qty,
  exrp_qty_total AS exrp_qty,
  owed_qty_total AS owed_qty,
  Switch (
    (acbal_total - pkup_qty_total) <= 0
    AND difm_qty_total > 0,
    "red",
    (acbal_total - pkup_qty_total) <= 0,
    "green",
    (
      (
        (rfi_qty_total) / IIf(
          (acbal_total - pkup_qty_total) = 0,
          1,
          (acbal_total - pkup_qty_total)
        )
      ) * 100
    ) < 33.4,
    "red",
    (
      (
        (rfi_qty_total) / IIf(
          (acbal_total - pkup_qty_total) = 0,
          1,
          (acbal_total - pkup_qty_total)
        )
      ) * 100
    ) <= 67,
    "yellow",
    True,
    "green"
  ) AS buffer_status,
  Round(
    (rfi_qty_total) / IIf(
      (acbal_total - pkup_qty_total) = 0,
      1,
      (acbal_total - pkup_qty_total)
    ) * 100,
    2
  ) AS supply_percent
FROM
  Grouped_inventory;
