INSERT INTO Tbl_SelectedADHOCs (
  SortOrder, ADHOCName, ADHOCPath, SavePath,
  SaveFileType, ExeWith
)
SELECT
  ConfigTbl_ADHOCEngine.SortOrder,
  ConfigTbl_ADHOCEngine.ADHOCName,
  ConfigTbl_ADHOCEngine.ADHOCPath,
  ConfigTbl_ADHOCEngine.SavePath,
  ConfigTbl_ADHOCEngine.SaveFileType,
  ConfigTbl_ADHOCEngine.ExeWith
FROM
  ConfigTbl_ADHOCEngine
WHERE
  (
    (
      (ConfigTbl_ADHOCEngine.ExeWith) = "NALCOMIS ADHOC"
    )
    AND (
      (ConfigTbl_ADHOCEngine.RunADHOC) = True
    )
  );
