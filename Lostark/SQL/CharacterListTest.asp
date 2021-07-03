Insert into Member
(ID, Password, ClientCode, ClientLevel, NickName, OutDate, State, LastDate, LastIP)
Values
('dbo', 'dbo', 'User001', 'User', '테스트', Null, 'Y', GetDate(), '192.168.0.1')

Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '혜약', '마법사(여)', '바드', '1440.00', 'Y', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/bard_big.png')
Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '교참', '암살자(여)', '블레이드', '1406.33', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/blade_big.png')

Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '통육', '무도가(여)', '배틀마스터', '1355.00', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/battlemaster_big.png')
Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '방쇄', '무도가(여)', '인파이터', '1325.00', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/infighter_big.png')



Select * From Member
Select * From CharacterList

Select
	CL.CharacterNickName
	, CL.CharacterDivision
	, CL.CharacterJob
	, CL.CharacterLevel
	, CL.CharacterIsMain
	/*
		, 대표캐릭터 사진, 대표캐릭터 닉네임, 대표캐릭터 레벨, 대표캐릭터 직업
	*/
From Member M
	Left Outer Join CharacterList CL
	On M.ID = CL.ID 
Where 1=1
And M.NickName = 'RHBY'

    Select 
        M. NickName 
	      , CL.CharacterNickName 
	      , CL.CharacterDivision 
	      , CL.CharacterJob 
	      , CL.CharacterLevel 
		  , CL.CharacterImgSrc
    From Member M 
    Left Outer Join CharacterList CL 
    On M.ID = CL.ID  
    Where 1=1 
    And CL.CharacterIsMain = 'Y' 

    And M.ID = 'divashuthron'

	-- 대표 캐릭터 변경 시 내 정보 이미지 변경 테스트
	Begin Tran
	Select * From CharacterList

	Update CharacterList
		Set CharacterIsMain = 'Y'
	Where CharacterNickName = '교참'

	Update CharacterList
		Set CharacterIsMain = 'N'
	Where CharacterNickName = '혜약'

	Select * From CharacterList
	Rollback Tran