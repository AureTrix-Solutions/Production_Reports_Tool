SELECT
  Requisitions.mcn,
  Requisitions.cage,
  Requisitions.cog,
  Requisitions.cur_mcn,
  Requisitions.cur_niin,
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
  Requisitions.ord_dttm
FROM
  Requisitions
WHERE
  Requisitions.lsc NOT LIKE "c*"
  AND Requisitions.lsc NOT IN (
    'ossuf', 'rcanc', 'xmsuf', 'ignor'
  );
