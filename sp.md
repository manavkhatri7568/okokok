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
