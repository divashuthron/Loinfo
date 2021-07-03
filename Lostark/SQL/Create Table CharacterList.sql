USE [Loinfo]
GO

/****** Object:  Table [dbo].[Member]    Script Date: 2021-07-03 ¿ÀÈÄ 7:38:43 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CharacterList](
	[IDX] [int] IDENTITY(1,1) NOT NULL,
	[ID] [varchar](25) NOT NULL,
	[CharacterNickName] [varchar](25) NOT NULL,
	[CharacterDivision] [varchar](25) NULL,
	[CharacterJob] [varchar](50) NULL,
	[CharacterLevel] [varchar](50) NULL,
	[CharacterIsMain] [varchar](25) NULL,
	[CharacterImgSrc] [varchar](2000) NULL,
	[InsertTime] [varchar](30) NULL
)
GO

ALTER TABLE [dbo].[CharacterList] ADD  CONSTRAINT [DF_CharacterList_InsertTime]  DEFAULT (getdate()) FOR [InsertTime]
GO


