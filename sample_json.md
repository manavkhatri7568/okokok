[
  {
    "match_id": "030100000015",
    "RecordID": "492",
    "MessageID": "409",
    "UseCaseID": "32",
    "match_type": "MATCH",
    "fields": [
      {
        "field": "Counterparty reference(G)",
        "lhs": "BCS_7015981",
        "rhs": "20930557",
        "match": false
      },
      {
        "field": "Counterpartyname",
        "lhs": "SCB HONGKON*HKG",
        "rhs": "CP NAME (590687)",
        "match": false
      },
      {
        "field": "Counterpartyname",
        "lhs": "SCB HONGKON*HKG",
        "rhs": "CP NAME (590687)",
        "match": false
      },
      {
        "field": "Notional(G)",
        "lhs": "702000",
        "rhs": "702100",
        "match": false
      },
      {
        "field": "NomuraEntity",
        "lhs": "NIP",
        "rhs": "NIP",
        "match": true
      },
      {
        "field": "Product",
        "lhs": "TOTAL RETURN SWAP",
        "rhs": "UNKNOWN RP",
        "match": false
      },
      {
        "field": "Currency(G)",
        "lhs": "HKD",
        "rhs": "HKD",
        "match": true
      },
      {
        "field": "Direction(G)",
        "lhs": "RECEIVE",
        "rhs": "RECEIVE",
        "match": true
      },
      {
        "field": "TradeDate",
        "lhs": "",
        "rhs": "10-07-2026",
        "match": false
      },
      {
        "field": "ValueDate(M)",
        "lhs": "10-07-2026",
        "rhs": "10-07-2026",
        "match": true
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "HSBCHKHHHKH",
        "match": false
      },
      {
        "field": "SSI intermediary bank",
        "lhs": "",
        "rhs": "HONG KONG + SHANGHAI BK CORPORATION",
        "match": false
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "",
        "match": false
      },
      {
        "field": "SSIBankName",
        "lhs": "HSBCHKHHHKH",
        "rhs": "",
        "match": false
      },
      {
        "field": "SSIBankName",
        "lhs": "HSBCHKHHHKH",
        "rhs": "",
        "match": false
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "",
        "match": false
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "NOMAGB2L",
        "match": false
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "BENEFICIARY NAME (3783286)",
        "match": false
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "472766",
        "match": false
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "HSBCHKHHHKH",
        "match": false
      },
      {
        "field": "SSI intermediary bank",
        "lhs": "",
        "rhs": "",
        "match": true
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "",
        "match": false
      },
      {
        "field": "SSI intermediary bank",
        "lhs": "",
        "rhs": "",
        "match": true
      },
      {
        "field": "SSIBankName",
        "lhs": "HSBCHKHHHKH",
        "rhs": "",
        "match": false
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "",
        "match": false
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "CIBBMYKL",
        "match": false
      },
      {
        "field": "SSI beneficiary name (BIC)",
        "lhs": "NOMAGB2LXXX",
        "rhs": "BENEFICIARY NAME (65398)",
        "match": false
      },
      {
        "field": "Account Number",
        "lhs": "502363856292.000000",
        "rhs": "1157146",
        "match": false
      }



code for where this json came and we have to change it so i can take rime flat file output and convert to json and store it
import pandas as pd
import json
import sys
from pathlib import Path


def parse_recon_file(file_path: str, delimiter: str = '\t') -> list[dict]:
    df = pd.read_csv(file_path, sep=delimiter, dtype=str, keep_default_na=False)
    df.columns = df.columns.str.strip()

    match_id_col = 'MATCH ID'
    source_col   = 'SOURCE SYSTEM'

    meta_columns = {
        match_id_col, source_col,
        'REC', 'INSTANCE NO', 'BUSINESS DATE',
        'BREAK COLUMN COUNT', 'BREAK COLUMNS',
        'MATCH PASS', 'MATCH TYPE',
    }

    top_level_unique = {'RecordID', 'MessageID', 'UseCaseID'}

    compare_columns = [
        c for c in df.columns
        if c not in meta_columns and c not in top_level_unique
    ]

    results = []

    for match_id, group in df.groupby(match_id_col, sort=False):
        sources = group[source_col].unique().tolist()

        lhs_src = sources[0] if len(sources) > 0 else None
        rhs_src = sources[1] if len(sources) > 1 else None

        def first_row(src):
            rows = group[group[source_col] == src]
            return rows.iloc[0] if not rows.empty else None

        lhs_row = first_row(lhs_src) if lhs_src else None
        rhs_row = first_row(rhs_src) if rhs_src else None

        def get_meta(col):
            for src in sources:
                row = first_row(src)
                if row is not None:
                    val = row.get(col, '').strip()
                    if val:
                        return val
            return ''

        first_row_any = group.iloc[0]
        match_type = first_row_any.get('MATCH TYPE', '')

        # ✅ fields is now a LIST of {field, lhs, rhs, match} objects
        fields = []
        for col in compare_columns:
            lhs_val = lhs_row.get(col, '').strip() if lhs_row is not None else ''
            rhs_val = rhs_row.get(col, '').strip() if rhs_row is not None else ''

            if lhs_row is not None and rhs_row is not None:
                matched = lhs_val.upper() == rhs_val.upper()
            else:
                matched = None

            fields.append({
                "field": strip_field_suffix(col),
                "lhs":   lhs_val,
                "rhs":   rhs_val,
                "match": matched,
            })

        result = {
            "match_id":   match_id,
            "RecordID":   get_meta('RecordID'),
            "MessageID":  get_meta('MessageID'),
            "UseCaseID":  get_meta('UseCaseID'),
            "match_type": match_type,
            "fields":     fields,          # ← list, not dict
        }

        results.append(result)

    return results

def strip_field_suffix(field_name: str) -> str:
    """Remove underscore and everything after it from a field name."""
    return field_name.split('_', 1)[0]

def filter_breaks_only(results: list[dict]) -> list[dict]:
    """Filter to only show fields where match is False."""
    filtered = []
    for r in results:
        breaks = [f for f in r['fields'] if f.get('match') is False]
        if breaks:
            filtered.append({**r, 'fields': breaks})
    return filtered


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python recon_compare.py <file_path> [--breaks-only]")
        sys.exit(1)

    file_path   = sys.argv[1]
    breaks_only = '--breaks-only' in sys.argv

    results = parse_recon_file(file_path)

    if breaks_only:
        results = filter_breaks_only(results)

    output = json.dumps(results, indent=2, ensure_ascii=False)
    print(output)

    out_path = Path(file_path).stem + '_compeee.json'
    with open(out_path, 'w') as f:
        f.write(output)
    print(f"\nOutput written to: {out_path}")





    then below is the sp script of spMatchAttemptInsert from db/programmability/Stored Procedure/dbo.spMatchAttemptInsert  


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[spMatchAttemptInsert]
    @RecordID           BIGINT,
    @MatchTier          VARCHAR(20),
    -- EXACT | FUZZY | COUNTERPARTY_LESS
    -- also written to MatchResult
    @CandidatesJSON     NVARCHAR(MAX),
    -- JSON array; see shape above
    @OutAttemptID       BIGINT      OUTPUT,
    @ReturnCode         INT         OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET @OutAttemptID = 0;
    SET @ReturnCode   = 0;

    -- -----------------------------------------------
    -- 1. Validate record exists
    -- -----------------------------------------------
    IF NOT EXISTS(
        SELECT 1
    FROM [dbo].[tblRecord] WITH (NOLOCK)
    WHERE  [RecordID] = @RecordID
    )
    BEGIN
        SET @ReturnCode = 1;
        RETURN;
    END

    -- -----------------------------------------------
    -- 2. Validate MatchTier value (mirrors table CHECK constraint)
    -- -----------------------------------------------
    IF @MatchTier NOT IN ('EXACT', 'FUZZY', 'COUNTERPARTY_LESS')
    BEGIN
        SET @ReturnCode = 2;
        RETURN;
    END

    -- -----------------------------------------------
    -- 3. Validate candidates JSON is not empty
    -- -----------------------------------------------
    IF @CandidatesJSON IS NULL OR LEN(LTRIM(RTRIM(@CandidatesJSON))) = 0
    BEGIN
        SET @ReturnCode = 3;
        RETURN;
    END

    BEGIN TRANSACTION;
    BEGIN TRY

        -- -----------------------------------------------
        -- 4. Insert one tblMatchAttempt row
        -- -----------------------------------------------
        INSERT INTO [dbo].[tblMatchAttempt]
        (
        [RecordID],
        [AttemptSequence],
        [MatchTier],
        [MatchResult],
        [TargetSystem],
        [ExternalRef],
        [OverallScore],
        [ThresholdUsed],
        [MatchMetadata],
        [AttemptedBy],
        [AttemptedAt],
        [CreatedByID],
        [CreatedDate],
        [UpdatedByID]
        )
    VALUES
        (
            @RecordID,
            1, -- AttemptSequence hardcoded
            @MatchTier,
            @MatchTier, -- MatchResult = MatchTier (same value)
            'PCM', -- TargetSystem hardcoded
            NULL, -- ExternalRef blank
            NULL, -- OverallScore – not supplied at insert time
            NULL, -- ThresholdUsed – not supplied at insert time
            NULL, -- MatchMetadata – not supplied at insert time
            'SYSTEM', -- AttemptedBy hardcoded
            GETDATE(),
            0, -- CreatedByID hardcoded
            GETDATE(),
            0               -- UpdatedByID hardcoded
        );

        SET @OutAttemptID = SCOPE_IDENTITY();

        -- -----------------------------------------------
        -- 5. Insert one tblMatchAttemptCandidate row per element in @CandidatesJSON
        -- -----------------------------------------------
        INSERT INTO [dbo].[tblMatchAttemptCandidate]
        (
        [AttemptID],
        [CandidateRank],
        [TargetSystem],
        [ExternalRef],
        [MatchedRecordID],
        [CandidateScore],
        [FieldDeltasJSON],
        [IsSelected],
        [SelectedByID],
        [SelectedAt],
        [CreatedByID],
        [CreatedDate],
        [UpdatedByID]
        )
    SELECT
        @OutAttemptID,
        c.[CandidateRank],
        'PCM', -- TargetSystem mirrors parent row
        c.[ExternalRef],
        NULL, -- MatchedRecordID – Record-vs-System mode; not used here
        NULL, -- CandidateScore – empty for now
        c.[FieldDeltasJSON],
        0, -- IsSelected default
        NULL, -- SelectedByID default
        NULL, -- SelectedAt default
        0, -- CreatedByID hardcoded
        GETDATE(),
        0
    -- UpdatedByID hardcoded
    FROM OPENJSON(@CandidatesJSON)
        WITH (
            [CandidateRank]     TINYINT         '$.candidate_rank',
            [ExternalRef]       VARCHAR(200)    '$.external_ref',
            [FieldDeltasJSON]   NVARCHAR(MAX)   '$.field_deltas_json'
        ) c;

        -- -----------------------------------------------
        -- 6. Audit event
        -- -----------------------------------------------
        INSERT INTO [dbo].[tblAuditEvent]
        (
        [EventType],
        [EntityType],
        [EntityID],
        [UserID],
        [Metadata],
        [CreatedByID],
        [CreatedDate],
        [UpdatedByID]
        )
    VALUES
        (
            'MATCH_ATTEMPT_CREATED',
            'RECORD',
            @RecordID,
            0,
            N'{"attempt_id":'  + CAST(@OutAttemptID AS NVARCHAR(20))
              + N',"match_tier":"' + @MatchTier + N'"'
              + N',"target_system":"PCM"}',
            0,
            GETDATE(),
            0
        );

        COMMIT TRANSACTION;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        SET @ReturnCode   = -1;
        SET @OutAttemptID = 0;
        THROW;
    END CATCH;

END
GO




so now make the code like that it convert the rime output to json like i share and then a python script which takes that json and give sp script which we can use to push data to the table

and one more sample 
Sample insert script is here - 
 
-- =============================================
-- Test : spMatchAttemptInsert
-- RecordID  = 476
-- MatchTier = 'COUNTERPARTY_LESS'
-- Candidates = 3
-- =============================================
 
DECLARE @AttemptID  BIGINT;
DECLARE @RC         INT;
 
EXEC [dbo].[spMatchAttemptInsert]
    @RecordID       = 476,
    @MatchTier      = 'COUNTERPARTY_LESS',
    @CandidatesJSON = N'[
        {
            "candidate_rank": 1,
            "external_ref": "T-100001",
            "field_deltas_json": "[{\"field\":\"trade_date\",\"lhs\":\"2026-08-01\",\"rhs\":\"2026-08-01\",\"match\":true},{\"field\":\"amount\",\"lhs\":\"500000\",\"rhs\":\"500000\",\"match\":true},{\"field\":\"counterparty\",\"lhs\":\"BANK_A\",\"rhs\":null,\"match\":false}]"
        },
        {
            "candidate_rank": 2,
            "external_ref": "T-100002",
            "field_deltas_json": "[{\"field\":\"trade_date\",\"lhs\":\"2026-08-01\",\"rhs\":\"2026-08-02\",\"match\":false},{\"field\":\"amount\",\"lhs\":\"500000\",\"rhs\":\"500000\",\"match\":true},{\"field\":\"counterparty\",\"lhs\":\"BANK_A\",\"rhs\":null,\"match\":false}]"
        },
        {
            "candidate_rank": 3,
            "external_ref": "T-100003",
            "field_deltas_json": "[{\"field\":\"trade_date\",\"lhs\":\"2026-08-01\",\"rhs\":\"2026-08-03\",\"match\":false},{\"field\":\"amount\",\"lhs\":\"500000\",\"rhs\":\"490000\",\"match\":false},{\"field\":\"counterparty\",\"lhs\":\"BANK_A\",\"rhs\":null,\"match\":false}]"
        }
    ]',
    @OutAttemptID   = @AttemptID OUTPUT,
    @ReturnCode     = @RC        OUTPUT;
 
-- ── Result summary ────────────────────────────────────────────────────────
SELECT
    @RC        AS ReturnCode,           -- expect 0
    @AttemptID AS OutAttemptID;         -- expect a new BIGINT identity value
 
 
--UPDATE dbo.tblLkpMatchResult
--SET Code = 'COUNTERPARTY_LESS'
--WHERE Code = 'COUNTERPARTY_LESS_CANDIDATE'

    ]
  }
]
