UPDATE
  Growler_ddsn_list
SET
  ddsn = Replace([ddsn], "-", "")
WHERE
  ddsn LIKE "*-*";
