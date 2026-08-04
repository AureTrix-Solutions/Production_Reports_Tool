UPDATE
  p8_ddsn_list
SET
  ddsn = Replace([ddsn], "-", "")
WHERE
  ddsn LIKE "*-*";
