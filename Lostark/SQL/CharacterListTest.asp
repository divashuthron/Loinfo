Insert into Member
(ID, Password, ClientCode, ClientLevel, NickName, OutDate, State, LastDate, LastIP)
Values
('dbo', 'dbo', 'User001', 'User', '�׽�Ʈ', Null, 'Y', GetDate(), '192.168.0.1')

Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '����', '������(��)', '�ٵ�', '1440.00', 'Y', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/bard_big.png')
Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '����', '�ϻ���(��)', '���̵�', '1406.33', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/blade_big.png')

Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '����', '������(��)', '��Ʋ������', '1355.00', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/battlemaster_big.png')
Insert Into CharacterList
(ID, CharacterNickName, CharacterDivision, CharacterJob, CharacterLevel, CharacterIsMain, CharacterImgSrc)
Values
('divashuthron', '���', '������(��)', '��������', '1325.00', 'N', 'https://cdn-lostark.game.onstove.com/2018/obt/assets/images/common/thumb/infighter_big.png')



Select * From Member
Select * From CharacterList

Select
	CL.CharacterNickName
	, CL.CharacterDivision
	, CL.CharacterJob
	, CL.CharacterLevel
	, CL.CharacterIsMain
	/*
		, ��ǥĳ���� ����, ��ǥĳ���� �г���, ��ǥĳ���� ����, ��ǥĳ���� ����
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

	-- ��ǥ ĳ���� ���� �� �� ���� �̹��� ���� �׽�Ʈ
	Begin Tran
	Select * From CharacterList

	Update CharacterList
		Set CharacterIsMain = 'Y'
	Where CharacterNickName = '����'

	Update CharacterList
		Set CharacterIsMain = 'N'
	Where CharacterNickName = '����'

	Select * From CharacterList
	Rollback Tran