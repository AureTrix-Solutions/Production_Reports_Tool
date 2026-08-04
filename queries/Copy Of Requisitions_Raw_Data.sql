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
  Requisitions.jcn,
  Requisitions.unit_price,
  Requisitions.net_price,
  Requisitions.ord_dttm,
  Requisitions.lsc_dttm,
  IIf(
    [cur_mcn] IS NULL,
    [wc_cd] & [jcn],
    [wc_cd] & Left([jcn], 9)
  ) AS cost_key
FROM
  Requisitions
  LEFT JOIN hof_niin ON Requisitions.fgc = hof_niin.fgc;
