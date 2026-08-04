SELECT
  Requisitions.mcn,
  Requisitions.cog,
  Requisitions.cur_niin,
  Requisitions.ddsn,
  Requisitions.ddsn_suf,
  Requisitions.lsc,
  Requisitions.ord_qty,
  Requisitions.part_no,
  Requisitions.nomen,
  Requisitions.fgc,
  Requisitions.rep_itm_ind,
  Requisitions.ord_dttm
FROM
  Requisitions
WHERE
  Requisitions.lsc NOT LIKE "c*"
  AND Requisitions.lsc NOT IN (
    'ossuf', 'rcanc', 'xmsuf', 'ignor'
  );
