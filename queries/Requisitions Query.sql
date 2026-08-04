UPDATE
  Requisitions
SET
  Requisitions.ord_dttm = Left(
    [ord_dttm_text],
    InStr([ord_dttm_text], " ") - 1
  ),
  Requisitions.lsc_dttm = Left(
    [lsc_dttm_text],
    InStr([lsc_dttm_text], " ") - 1
  );
