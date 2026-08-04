UPDATE
  Maf_Notes
SET
  Maf_Notes.note_dttm = Left(
    [note_dttm_text],
    InStr([note_dttm_text], " ") - 1
  );
