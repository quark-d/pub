```sql
CREATE PROCEDURE SyncStatusesFromA_Multiple
AS
BEGIN
    SET NOCOUNT ON;

    -- bテーブル
    MERGE b AS TARGET
    USING (SELECT wo, statusB FROM a) AS SOURCE
    ON TARGET.wo = SOURCE.wo
    WHEN MATCHED AND ISNULL(TARGET.statusB, '') <> ISNULL(SOURCE.statusB, '')
        THEN UPDATE SET TARGET.statusB = SOURCE.statusB
    WHEN NOT MATCHED BY TARGET
        THEN INSERT (wo, statusB) VALUES (SOURCE.wo, SOURCE.statusB);

    -- cテーブル
    MERGE c AS TARGET
    USING (SELECT wo, statusC FROM a) AS SOURCE
    ON TARGET.wo = SOURCE.wo
    WHEN MATCHED AND ISNULL(TARGET.statusC, '') <> ISNULL(SOURCE.statusC, '')
        THEN UPDATE SET TARGET.statusC = SOURCE.statusC
    WHEN NOT MATCHED BY TARGET
        THEN INSERT (wo, statusC) VALUES (SOURCE.wo, SOURCE.statusC);

    -- dテーブル
    MERGE d AS TARGET
    USING (SELECT wo, statusD FROM a) AS SOURCE
    ON TARGET.wo = SOURCE.wo
    WHEN MATCHED AND ISNULL(TARGET.statusD, '') <> ISNULL(SOURCE.statusD, '')
        THEN UPDATE SET TARGET.statusD = SOURCE.statusD
    WHEN NOT MATCHED BY TARGET
        THEN INSERT (wo, statusD) VALUES (SOURCE.wo, SOURCE.statusD);
END
```
