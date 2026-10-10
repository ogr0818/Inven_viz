-- schema 版本 1；由 manage_summary.py 的 SQLAlchemy 模型產生。

PRAGMA foreign_keys = ON;

PRAGMA user_version = 1;

CREATE TABLE drug_codes (
	drug_code TEXT NOT NULL, 
	PRIMARY KEY (drug_code), 
	CONSTRAINT ck_drug_code_nonempty CHECK (length(trim(drug_code)) > 0)
);

CREATE TABLE drug_master (
	drug_code TEXT NOT NULL, 
	drug_name TEXT NOT NULL, 
	drug_type TEXT NOT NULL, 
	PRIMARY KEY (drug_code), 
	FOREIGN KEY(drug_code) REFERENCES drug_codes (drug_code) ON DELETE RESTRICT
);

CREATE TABLE monthly_summary (
	drug_code TEXT NOT NULL, 
	roc_year INTEGER NOT NULL, 
	roc_month INTEGER NOT NULL, 
	drug_name TEXT NOT NULL, 
	drug_type TEXT NOT NULL, 
	issued_quantity INTEGER NOT NULL, 
	inpatient_usage INTEGER NOT NULL, 
	stock_quantity INTEGER NOT NULL, 
	PRIMARY KEY (drug_code, roc_year, roc_month), 
	CONSTRAINT ck_monthly_roc_year CHECK (roc_year > 0), 
	CONSTRAINT ck_monthly_roc_month CHECK (roc_month BETWEEN 1 AND 12), 
	CONSTRAINT ck_monthly_issued_positive CHECK (issued_quantity > 0), 
	FOREIGN KEY(drug_code) REFERENCES drug_codes (drug_code) ON DELETE RESTRICT
);

CREATE INDEX ix_monthly_summary_period ON monthly_summary (roc_year, roc_month);
