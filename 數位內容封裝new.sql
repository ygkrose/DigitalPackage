
CREATE	TABLE	UserTempCaDetail	/*UTCAD �ϥΪ̸�ñ����O����*/
	(
	UseOrgSeqNO	NUMERIC(2,0) NOT NULL,	/* �����N�X */
	DocNO	CHAR(12) NOT NULL,	/* ����帹 */
	RpsDepID	CHAR(10) NOT NULL,	/* �n�O��� */
	RpsDepNam	NVARCHAR(40) NOT NULL,	/*  */
	RpsDivID	CHAR(10) NOT NULL,	/*  */
	RpsDivNam	NVARCHAR(40) NOT NULL,	/*  */
	RpsUserID	CHAR(16) NOT NULL,	/* �n�O�H�� */
	RpsUserNam	NVARCHAR(20) NOT NULL,	/*  */
	RpsTime	VARCHAR(18) NOT NULL,	/* �n�O�ɶ� */	
	RemDepID	CHAR(10) NOT NULL,	/* ��ñ��� */
	RemDepNam	NVARCHAR(40) NOT NULL,	/*  */
	RemDivID	CHAR(10) NOT NULL,	/*  */
	RemDivNam	NVARCHAR(40) NOT NULL,	/*  */
	RemUserID	CHAR(16) NOT NULL,	/* ��ñ�H�� */
	RemUserNam	NVARCHAR(20) NOT NULL,	/*  */
	RemTime	VARCHAR(18) NOT NULL,	/* ��ñ�ɶ� */
	FlowID		NVARCHAR(150) NOT NULL,	/* ñ���I�w�q Id �ݩʭ�2 :ñ���IID�Ǹ� */
	RemReason	NVARCHAR(20) NOT NULL,	/* ��ñ�������O */
	RsvFieldA	NVARCHAR(20) NOT NULL,	/* �w�d��� */
	RsvFieldB	NVARCHAR(20) NOT NULL,	/* �w�d��� */
	CreTime	VARCHAR(18) NOT NULL,	/* �إ߮ɶ� */
	CONSTRAINT UTCAD_PKEY PRIMARY KEY (UseOrgSeqNO,DocNO,FlowID)
	)
;
