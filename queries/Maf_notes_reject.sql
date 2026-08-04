SELECT
  m1.mcn,
  m1.subj,
  m1.maf_note,
  m1.note_dttm
FROM
  maf_notes AS m1
  INNER JOIN (
    SELECT
      mcn,
      MAX(note_dttm) AS max_note_dttm
    FROM
      maf_notes
    GROUP BY
      mcn
  ) AS m2 ON (m1.note_dttm = m2.max_note_dttm)
  AND (m1.mcn = m2.mcn)
WHERE
  m1.subj LIKE "*reject*"
ORDER BY
  m1.note_dttm DESC;
