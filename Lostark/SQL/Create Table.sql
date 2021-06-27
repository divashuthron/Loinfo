IF OBJECT_ID('LoaRoom.dbo.Member', 'U') IS NOT NULL
  DROP TABLE Member
IF OBJECT_ID('LoaRoom.dbo.CharacterInformation', 'U') IS NOT NULL
  DROP TABLE CharacterInformation
IF OBJECT_ID('LoaRoom.dbo.ContentSchedule', 'U') IS NOT NULL
  DROP TABLE ContentSchedule
IF OBJECT_ID('LoaRoom.dbo.Party', 'U') IS NOT NULL
  DROP TABLE Party
IF OBJECT_ID('LoaRoom.dbo.PartyMemo', 'U') IS NOT NULL
  DROP TABLE PartyMemo
IF OBJECT_ID('LoaRoom.dbo.PartyWaitList', 'U') IS NOT NULL
  DROP TABLE PartyWaitList

CREATE TABLE dbo.Member (
	IDX int IDENTITY(1,1) NOT NULL,
	ID varchar(50) NOT NULL,
	Password varchar(50) NULL,
	Email varchar(50) NULL,
	NickName varchar(50) NULL,
	InsertTime DateTime NOT NULL,
 CONSTRAINT PK_Member PRIMARY KEY CLUSTERED 
(
	ID ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE dbo.Member ADD  DEFAULT (getdate()) FOR InsertTime
GO

CREATE TABLE dbo.CharacterInformation (
	IDX int IDENTITY(1,1) NOT NULL,
	ID varchar(50) NOT NULL,
	CharacterName varchar(20) NOT NULL,
	Class varchar(20) NULL,
	Division1 varchar(20) NULL,
	Division2 varchar(20) NULL,
	Division3 varchar(20) NULL,
	Division4 varchar(20) NULL,
	ItemLevel varchar(20) NULL,
	IsMaster varchar(20) NULL,
	SettingName1 varchar(20) NULL,
	SettingLevel1 varchar(20) NULL,
	SettingName2 varchar(20) NULL,
	SettingLevel2 varchar(20) NULL,
	InsertTime DateTime NOT NULL,
	InsertUserID varchar(50) NULL,
 CONSTRAINT PK_CharacterInformation PRIMARY KEY CLUSTERED 
(
	CharacterName ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE dbo.CharacterInformation ADD  DEFAULT (getdate()) FOR InsertTime
GO

CREATE TABLE dbo.ContentSchedule (
	IDX int IDENTITY(1,1) NOT NULL,
	ID varchar(50) NOT NULL,
	CharacterName varchar(20) NOT NULL,
	Argos varchar(10) NULL,
	ValtanNormal varchar(10) NULL,
	ValtanHard varchar(10) NULL,
	BiackissNormal varchar(10) NULL,
	BiackissHard varchar(10) NULL,
	KoukuSatonRehearsal varchar(10) NULL,
	KoukuSatonName varchar(10) NULL,
	ResetTime DateTime NOT NULL,
 CONSTRAINT PK_ContentSchedule PRIMARY KEY CLUSTERED 
(
	CharacterName ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

CREATE TABLE dbo.Party (
	IDX int IDENTITY(1,1) NOT NULL,
	PartyID varchar(50) NOT NULL,
	PartyName varchar(2000) NULL,
	RaidName varchar(50) NULL,
	RaidDifficult varchar(50) NULL,
	PartyGoal varchar(10) NULL,
	PartyDate DateTime NULL,
	PartyStatus varchar(10) NULL,

	MasterName varchar(20) NULL,
	MemberName1 varchar(20) NULL,
	MemberName2 varchar(20) NULL,
	MemberName3 varchar(20) NULL,
	MemberName4 varchar(20) NULL,
	MemberName5 varchar(20) NULL,
	MemberName6 varchar(20) NULL,
	MemberName7 varchar(20) NULL,

	MasterCheck varchar(10) NULL,
	MemberCheck1 varchar(10) NULL,
	MemberCheck2 varchar(10) NULL,
	MemberCheck3 varchar(10) NULL,
	MemberCheck4 varchar(10) NULL,
	MemberCheck5 varchar(10) NULL,
	MemberCheck6 varchar(10) NULL,
	MemberCheck7 varchar(10) NULL,

	InsertTime DateTime NOT NULL,
	InsertUserID varchar(50) NULL,
	UpdateTime DateTime NOT NULL,
	UpdateUserID varchar(50) NULL,
 CONSTRAINT PK_Party PRIMARY KEY CLUSTERED 
(
	PartyID ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

ALTER TABLE dbo.Party ADD  DEFAULT (getdate()) FOR InsertTime
GO

CREATE TABLE dbo.PartyMemo (
	IDX int IDENTITY(1,1) NOT NULL,
	PartyID varchar(50) NOT NULL,
	CharacterName varchar(20) NOT NULL,
	Memo varchar(2000) NULL,
	InsertTime DateTime NOT NULL,
	InsertUserID varchar(50) NULL,
	UpdateTime DateTime NOT NULL,
	UpdateUserID varchar(50) NULL
)
GO

ALTER TABLE dbo.PartyMemo ADD  DEFAULT (getdate()) FOR InsertTime
GO

CREATE TABLE dbo.PartyWaitList (
	IDX int IDENTITY(1,1) NOT NULL,
	PartyID varchar(50) NOT NULL,
	CharacterName varchar(20) NOT NULL,
	InsertTime DateTime NOT NULL,
	InsertUserID varchar(50) NULL,
)
GO

ALTER TABLE dbo.PartyWaitList ADD  DEFAULT (getdate()) FOR InsertTime
GO