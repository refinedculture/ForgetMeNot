
USE [ForgetMeNot]
GO

/****** Object:  Table [dbo].[ForgetMeNot]    Script Date: 06/17/2014 21:15:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ForgetMeNot](
	[Application_ID] [int] IDENTITY(1,1) NOT NULL,
	[Display_Name] [varchar](max) NULL,
	[Application_Name] [varchar](max) NULL,
	[Application_URL] [varchar](max) NULL,
	[Login_URL] [varchar](max) NULL,
	[Username] [varchar](max) NULL,
	[Email] [varchar](max) NULL,
	[Password] [varchar](max) NULL,
	[Software_Key] [varchar](max) NULL,
	[CD_Key] [varchar](max) NULL,
	[General_Key_1] [varchar](max) NULL,
	[General_Key_2] [varchar](max) NULL,
	[General_Key_3] [varchar](max) NULL,
	[General_Key_4] [varchar](max) NULL,
	[General_Key_5] [varchar](max) NULL,
	[Has_Credit_Card_Info] [bit] NOT NULL,
	[Has_Address_Info] [bit] NOT NULL,
	[Has_Phone_Number] [bit] NOT NULL,
	[Is_Auto_Renewal] [bit] NOT NULL,
	[Notes] [varchar](max) NULL,
 CONSTRAINT [PK_ForgetMeNot] PRIMARY KEY CLUSTERED 
(
	[Application_ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO
