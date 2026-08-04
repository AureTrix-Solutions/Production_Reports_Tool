SELECT
  combined.fgc AS mom_fgc,
  combined.tec AS mom_tec,
  combined.niin AS mom_niin,
  combined.e_serno AS mom_serno,
  combined.mcn AS mom_mcn,
  combined.current_status AS mom_maf_status,
  combined.current_status_date AS mom_status_date,
  combined.act_take_cd AS mom_act_take_cd,
  combined.ty_maf_cd,
  IIf(
    combined.net_price = 0, combined.unit_price,
    combined.net_price
  ) AS mom_repair_cost,
  [Copy Of Requisitions_Raw_Data].cage,
  [Copy Of Requisitions_Raw_Data].cog,
  [Copy Of Requisitions_Raw_Data].cur_mcn,
  [Copy Of Requisitions_Raw_Data].cur_niin,
  [Copy Of Requisitions_Raw_Data].ddsn,
  [Copy Of Requisitions_Raw_Data].ddsn_suf,
  [Copy Of Requisitions_Raw_Data].lsc,
  [Copy Of Requisitions_Raw_Data].ord_qty,
  [Copy Of Requisitions_Raw_Data].part_no,
  [Copy Of Requisitions_Raw_Data].rep_mcn,
  combined.wc_cd,
  [Copy Of Requisitions_Raw_Data].nomen,
  [Copy Of Requisitions_Raw_Data].ui,
  [Copy Of Requisitions_Raw_Data].fgc,
  [Copy Of Requisitions_Raw_Data].rep_itm_ind,
  [Copy Of Requisitions_Raw_Data].buno_serno,
  [Copy Of Requisitions_Raw_Data].cur_ti_ddsn,
  [Copy Of Requisitions_Raw_Data].jcn,
  CCur (
    IIf(
      [copy of Requisitions_Raw_Data].net_price IS NULL,
      0,
      IIf(
        [copy of requisitions_raw_data].net_price = 0,
        [copy of Requisitions_Raw_Data].unit_price,
        [copy of Requisitions_Raw_Data].net_price
      ) * [copy of Requisitions_Raw_Data].ord_qty,
      )
  ) AS repair_part_cost,
  [Copy Of Requisitions_Raw_Data].ord_dttm,
  [Copy Of Requisitions_Raw_Data].lsc_dttm
FROM
  combined
  LEFT JOIN [Copy Of Requisitions_Raw_Data] ON combined.cost_key = [Copy Of Requisitions_Raw_Data].cost_key
WHERE
  (
    (
      (combined.fgc) IS NOT NULL
    )
    AND (
      (combined.niin) NOT LIKE "L*"
      AND (combined.niin) <> "00H60MN00"
    )
    AND (
      (combined.current_status) = "jc"
    )
    AND (
      (combined.act_take_cd) NOT IN (
        "1", "7", "8", "9", "d", "n", "4", "5",
        "2"
      )
    )
    AND (
      (combined.ty_maf_cd) NOT IN (
        "TD", "SC", "WR", "PC", "CC", "AD", "SW"
      )
    )
    AND combined.sys_rsn not like "*inventory*"
    AND (
      (
        [Copy Of Requisitions_Raw_Data].lsc
      ) NOT IN (
        "ignor", "cancl", "refer", "exrep",
        "xmsuf", "rcanc"
      )
      OR [copy Of Requisitions_Raw_Data].lsc IS NULL
    )
    AND (
      (combined.wc_cd) NOT IN ("050", "sas")
      AND combined.wc_cd IS NOT NULL
    )
  );
