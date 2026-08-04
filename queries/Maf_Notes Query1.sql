DELETE Maf_Notes.id,
Maf_Notes.mcn,
Maf_Notes.subj,
Maf_Notes.maf_note,
Maf_Notes.note_dttm,
Maf_Notes.orignr
FROM
  Maf_Notes
WHERE
  (
    (
      (Maf_Notes.mcn) Is Null
    )
  );
