#Requires AutoHotkey v2.0

;--------------------------------------------------------
; MyAssign made with AutoHotKey
;--------------------------------------------------------
; update history
;
; 2022/02/04 �o�b�N�A�b�v����Win10�p�ɏC���AIME��on/off���폜
;            AHK�̃o�[�W�����A�b�v�ɔ����Avk1Dsc07B �� vk1D �̂悤�ɃL�[�������ύX
; 2022/02/07 alt+k ���ꕔ�̊��Ő��������삵�Ȃ��������߁A�R�����g���C��
;            ���ϊ��{v�Ńo�[�W������\������悤�ɂ���
; 2022/06/28 Excel��̓�������P
; 2022/07/04 IME��On/Off�𕜊�
; 2022/07/06 Excel��̓�������P
; 2022/07/07 Excel�ŏ����Ȃ��y�[�X�g��ǉ�
; 2022/07/07 Excel��F1��F2�ɂ���
; 2023/03/09 SHIFT + CTRL + P�i�p�X���[�h�����j��VSCode��Ŗ������i�R�}���h�p���b�g�\���Əd�����邽�߁j
; 2023/03/28 �e�L�X�g�{�b�N�X����TAB�������y�[�X�g���悤�Ƃ��������܂������Ȃ�
; 2023/07/25 OneNote�ŏ㉺�L�[���������Ȃ����ɑΉ�
;            ���ϊ��{nm,.�Ń}�E�X�z�C�[���G�~�����[�V������ǉ�
; 2024/01/23 PowerPoint�ŁA�}�`�̏����ɑΉ�
; 2025/01/24 �Ђ炪��/�J�^�J�i�L�[�ł�IME��ON�ɂ���
; 2025/03/07 AutoHotKey v2.0�ɑΉ�
; 2025/03/10 �����exe�̂ݗL���Ȑݒ肪�����Ă��Ȃ������̂��C���A������excel�����O����ݒ肪��Ɍ����Ă��܂���肠��
; 2025/04/11 �܂�ɔ��p�p���ɂȂ��Ė߂��Ȃ��Ȃ�����C���A�뎚�C��
; 2025/04/14 Win�{�ϊ��AWin�{�J�^�J�i�ł����p�p���ɂȂ��Ă��܂������C��
;            ���������݁AShift�{�J�^�J�i�őS�p�J�^�J�i�ɂȂ��Ă��܂���肪����ł��Ă��Ȃ�
; 2025/05/08 ���ϊ��{�G�i���t���́j�̃G�N�Z�����O�������Ă��Ȃ������̂��C��
;            Shift�{Ctrl�{P�i�p�X���[�h�����j��VSCode���O�������Ă��Ȃ������̂��C��
;--------------------------------------------------------
; �L�[�A�T�C���ꗗ 
; ���ϊ�+v�@		�o�[�W�����\��
; ���ϊ�		IME �� OFF
; �ϊ�			IME �� ON
; �J�^�J�i		IME �� ON
; Ctrl+;		���t���́iyyyy/mm/dd�`���j���G�N�Z��������
; Ctrl+Shift+;		���t���́iyyyymmdd�`���j���G�N�Z��������
; Ctrl+Shift+V	�����Ȃ��y�[�X�g�@���G�N�Z���̂�
; Ctrl+1�i�t���L�[�j	�}�`�̏����@��PowerPoint�̂�
; 
; ��ALT+h		���J�[�\���ړ� ���V�t�g���p�őI��
; ���ϊ�+h		���J�[�\���ړ� ���V�t�g���p�őI��
; 
; ��ALT+j		���J�[�\���ړ� ���V�t�g���p�őI��
; ���ϊ�+j		���J�[�\���ړ� ���V�t�g���p�őI��
; 
; ��ALT+k		��J�[�\���ړ� ���V�t�g���p�őI��
; ���ϊ�+k		��J�[�\���ړ� ���V�t�g���p�őI��
; 
; ��ALT+l		�E�J�[�\���ړ� ���V�t�g���p�őI��
; ���ϊ�+l		�E�J�[�\���ړ� ���V�t�g���p�őI��
; 
; ��ALT+u		PageDown
; ���ϊ�+u		PageDown
; 
; ��ALT+i		PageUp
; ���ϊ�+i		PageUp
; 
; ��ALT+y		Home
; ���ϊ�+y		Home
; 
; ��ALT+o		End
; ���ϊ�+o		End
; 
; Shift+Ctrl+p		�p�X���[�h���� ��VSCode������
; 
; ���ϊ�+t		TAB�������y�[�X�g
; 
; ���ϊ�+m		�X�N���[���_�E��
; ���ϊ�+,		�X�N���[���A�b�v
; ���ϊ�+n		���X�N���[��
; ���ϊ�+.		�E�X�N���[��
;--------------------------------------------------------

#Include IMEv2.ahk

; ��v�ȃL�[�� HotKey �ɐݒ肵�A���������p�X�X���[����
*~a::
*~b::
*~c::
*~d::
*~e::
*~f::
*~g::
*~h::
*~i::
*~j::
*~k::
*~l::
*~m::
*~n::
*~o::
*~p::
*~q::
*~r::
*~s::
*~t::
*~u::
*~v::
*~w::
*~x::
*~y::
*~z::
*~1::
*~2::
*~3::
*~4::
*~5::
*~6::
*~7::
*~8::
*~9::
*~0::
*~F1::
*~F2::
*~F3::
*~F4::
*~F5::
*~F6::
*~F7::
*~F8::
*~F9::
*~F10::
*~F11::
*~F12::
*~`::
*~~::
*~!::
*~@::
*~#::
*~$::
*~%::
*~^::
*~&::
*~*::
*~(::
*~)::
*~-::
*~_::
*~=::
*~+::
*~[::
*~{::
*~]::
*~}::
*~\::
*~|::
*~;::
*~'::
*~"::
*~,::
*~<::
*~.::
*~>::
*~/::
*~?::
*~Esc::
*~Tab::
*~Space::
*~LAlt::
*~RAlt::
*~Left::
*~Right::
*~Up::
*~Down::
*~Enter::
*~PrintScreen::
*~Delete::
*~Home::
*~End::
*~PgUp::
*~PgDn::
;*~vk1D:: ;���ϊ�
;*~vk1C:: ;�ϊ�
{
	Return
}

;;******************************************************************************
;;  �e�X�g�p
;;******************************************************************************
;vkF2::
;{
;	MsgBox("test")
;	return	
;}


;******************************************************************************
;  �o�[�W�����\��
;******************************************************************************
;���ϊ�+v���o�[�W�����\��
vk1D & v::
{
	MsgBox ("MyAssign last update 2025/05/08")
	return
}

;******************************************************************************
;  IME��On/Off
;******************************************************************************
; ���ϊ� �� IME �� OFF
vk1D::
 {
	IME_SET(0)
	Return
}
; �ϊ��A�J�^�J�i �� IME �� ON
vk1C::
vkF2::
{
	IME_SET(1)
	Return
}
; Ctrl�{�J�^�J�i�AWin�{�J�^�J�i�AShift�{�J�^�J�i�AWin�{�ϊ��𖳌���
^vkF2::
#vkF2::
+vkF2::
#vk1C::
{
	Return
}


;******************************************************************************
;  ���t�̓���
;******************************************************************************
; Excel�ł̓�������O����
;#HotIf not WinActive("ahk_exe EXCEL.EXE") ; Excel�ȊO���A�N�e�B�u�ȂƂ�
#HotIf WinGetProcessName("A") != "EXCEL.EXE"
	;Ctrl+;�œ��t����(yyyy/mm/dd�`��)
	^vkBB::
	{
		A_Clipboard:=A_YYYY "/" A_MM "/" A_DD
		Send "^v"
		Return
	}
	;Ctrl+Shift+;�œ��t����(yyyymmdd�`��)
	^+vkBB::
	{
		A_Clipboard:=A_YYYY A_MM A_DD
		Send "^v"
		Return
	}
#HotIf

;******************************************************************************
;  Excel�݂̂ŗL���ȃL�[
;******************************************************************************
#HotIf WinActive("ahk_exe EXCEL.EXE") ; Excel���A�N�e�B�u�ȂƂ�
	; F1�w���v�𖳌���
	F1::
	{
		Return
	}	
	; Ctrl�{Shift�{V �ŏ����Ȃ��y�[�X�g
	^+v::
	{
		Send "{AppsKey}v"
		Return
	}
#HotIf


;******************************************************************************
;  PowerPoint�݂̂ŗL���ȃL�[
;******************************************************************************
#HotIf WinActive("ahk_exe POWERPNT.EXE")
	; Ctrl�{1�i�t���L�[�j �Ő}�`�̏���
	^1::
	{
		Send "{AppsKey}o{Enter}"
		Return
	}	
#HotIf

;******************************************************************************
;  vi�̃J�[�\���ړ��}�b�v
;******************************************************************************

;��ALT+h����
<!h::
{	if GetKeyState("shift", "P"){
		Send "+{Left}"
	}else{
		Send "{Left}"
	}
	return
}
;��ALT+j����
<!j::
{
	if GetKeyState("shift", "P"){
		Send "+{Down}"
	}else{
		Send "{Down}"
	}
	return
}
;��ALT+k������
<!k::
{
	If GetKeyState("shift", "P"){
		Send "+{Up}"
	}else{
		Send "{Up}"
	}
	return
}
;��ALT+l���E
<!l::
{
	If GetKeyState("shift", "P"){
		Send "+{Right}"
	}else{
		Send "{Right}"
	}
	return
}
;���ϊ�+h����
vk1D & h::
{
	if GetKeyState("shift", "P"){
		Send "+{Left}"
	}else{
		Send "{Left}"
	}
	return
}
;���ϊ�+j����
vk1D & j::
{
	if GetKeyState("shift", "P"){
		Send "+{Down}"
	}else{
		Send "{Down}"
	}
	return
}
;���ϊ�+k������
vk1D & k::
{
	If GetKeyState("shift", "P"){
		Send "+{Up}"
	}else{
		Send "{Up}"
	}
	return
}
;���ϊ�+l���E
vk1D & l::
{
	If GetKeyState("shift", "P"){
		Send "+{Right}"
	}else{
		Send "{Right}"
	}
	return
}
; left alt+yuio��PageDown/PageUp/Home/End
<!u::
{
	Send "{PgDn}"
	return
}
<!i::
{
	Send "{PgUp}"
	return
}
<!y::
{
	Send "{Home}"
	return
}
<!o::
{
	Send "{End}"
	return
}

; ���ϊ�+yuio��PageDown/PageUp/Home/End
vk1D & u::
{
	Send "{PgDn}"
	return
}
vk1D & i::
{
	Send "{PgUp}"
	return
}
vk1D & y::
{
	Send "{Home}"
	return
}
vk1D & o::
{
	Send "{End}"
	return
}

;******************************************************************************
; �p�X���[�h��������
;******************************************************************************
; VS Code �ł̓�������O����
;#HotIf !WinActive("ahk_exe Code.exe")
#HotIf WinGetProcessName("A") != "Code.exe"
	^+p::
	{
		Number := "23456789"
		Lowercase := "abcdefghjkmnpqrstuvwxyz"
		Uppercase := "ABCDEFGHJKMNPQRSTUVWXYZ"
		Mark := "!#$`%@?+-*;="
		Password := ""
		Loop 8
		{
			Start := Random(1, 4)
			If Start = 1
			{
				Type := Number
			}
			If Start = 2
			{
				Type := Lowercase
			}
			If Start = 3
			{
				Type := Uppercase
			}
			If Start = 4
			{
				Type := Mark
			}
			
			StringLen := StrLen(Type)
			Start := Random(1, StringLen)
			VCharacter := SubStr(Type, Start, 1)
			
			Password := Password VCharacter
		}
		A_Clipboard:=Password
		Send "^v"
		return
	}
#HotIf

;******************************************************************************
;  TAB�̓��́iTab�L�[�Ŏ��̃R���g���[���Ɉړ�����̂ł͂Ȃ��ATab�������y�[�X�g����j
;******************************************************************************
;���ϊ�+t
vk1D & t::
{
	A_Clipboard:=A_Tab
	Send "^v"
	return
}

;******************************************************************************
;  OneNote�݂̂̐ݒ�
;******************************************************************************
#HotIf WinActive("ahk_class Framework::CFrame")
    <!k::      DllCall("keybd_event", "int", 0x26, "int", 0, "int", 1, "int", 0) ;Up
    vk1D & k:: DllCall("keybd_event", "int", 0x26, "int", 0, "int", 1, "int", 0) ;Up
    <!j::      DllCall("keybd_event", "int", 0x28, "int", 0, "int", 1, "int", 0) ;Down
    vk1D & j:: DllCall("keybd_event", "int", 0x28, "int", 0, "int", 1, "int", 0) ;Down
#HotIf

;******************************************************************************
;  �}�E�X�z�C�[�� �G�~�����[�V����
;******************************************************************************
;���ϊ�+m���X�N���[���_�E��
vk1D & m::
{
	MouseClick "WheelDown"
	return
}
;���ϊ�+,���X�N���[���_�E��
vk1D & ,::
{
	MouseClick "WheelUp"
	return
}
;���ϊ�+n�����X�N���[��
vk1D & n::
{
	MouseClick "WheelLeft"
	return
}
;���ϊ�+.���E�X�N���[��
vk1D & .::
{
	MouseClick "WheelRight"
	return
}
