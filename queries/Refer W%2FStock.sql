SELECT
  Requisitions.mcn,
  Requisitions.cage,
  Requisitions.cog,
  Requisitions.cur_mcn,
  Requisitions.cur_niin,
  hof_niin.hof_niin,
  Requisitions.ddsn,
  Requisitions.ddsn_suf,
  Requisitions.lsc,
  Requisitions.ord_qty,
  Requisitions.part_no,
  Requisitions.rep_mcn,
  Requisitions.wc_cd,
  Requisitions.nomen,
  Requisitions.ui,
  Requisitions.fgc,
  Requisitions.rep_itm_ind,
  Requisitions.buno_serno,
  Requisitions.cur_ti_ddsn,
  Left([jcn], 3) AS org,
  Requisitions.jcn,
  Requisitions.unit_price,
  Requisitions.net_price,
  Requisitions.ord_dttm,
  grouped_inventory.rfi_qty_total AS rfi_qty
FROM
  (
    Requisitions
    INNER JOIN grouped_inventory ON requisitions.fgc = grouped_inventory.fgc
  )
  LEFT JOIN hof_niin ON Requisitions.fgc = hof_niin.fgc
WHERE
  LSC = "refer"
  AND rep_itm_ind = "y"
  AND rfi_qty_total > 0;
