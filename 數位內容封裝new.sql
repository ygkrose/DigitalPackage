
CREATE	TABLE	UserTempCaDetail	/*UTCAD 使用者補簽公文記錄表*/
	(
	UseOrgSeqNO	NUMERIC(2,0) NOT NULL,	/* 機關代碼 */
	DocNO	CHAR(12) NOT NULL,	/* 公文文號 */
	RpsDepID	CHAR(10) NOT NULL,	/* 登記單位 */
	RpsDepNam	NVARCHAR(40) NOT NULL,	/*  */
	RpsDivID	CHAR(10) NOT NULL,	/*  */
	RpsDivNam	NVARCHAR(40) NOT NULL,	/*  */
	RpsUserID	CHAR(16) NOT NULL,	/* 登記人員 */
	RpsUserNam	NVARCHAR(20) NOT NULL,	/*  */
	RpsTime	VARCHAR(18) NOT NULL,	/* 登記時間 */	
	RemDepID	CHAR(10) NOT NULL,	/* 補簽單位 */
	RemDepNam	NVARCHAR(40) NOT NULL,	/*  */
	RemDivID	CHAR(10) NOT NULL,	/*  */
	RemDivNam	NVARCHAR(40) NOT NULL,	/*  */
	RemUserID	CHAR(16) NOT NULL,	/* 補簽人員 */
	RemUserNam	NVARCHAR(20) NOT NULL,	/*  */
	RemTime	VARCHAR(18) NOT NULL,	/* 補簽時間 */
	FlowID		NVARCHAR(150) NOT NULL,	/* 簽核點定義 Id 屬性值2 :簽核點ID序號 */
	RemReason	NVARCHAR(20) NOT NULL,	/* 補簽說明註記 */
	RsvFieldA	NVARCHAR(20) NOT NULL,	/* 預留欄位 */
	RsvFieldB	NVARCHAR(20) NOT NULL,	/* 預留欄位 */
	CreTime	VARCHAR(18) NOT NULL,	/* 建立時間 */
	CONSTRAINT UTCAD_PKEY PRIMARY KEY (UseOrgSeqNO,DocNO,FlowID)
	)
;
