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

;******************************************************************************
;  �o�[�W�����\��
;******************************************************************************
;���ϊ�+v���o�[�W�����\��
vk1D & v::
{
	MsgBox ("MyAssign last update 2025/03/07")
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
; �ϊ� �� IME �� ON
vk1C::
vkF2::
{
	IME_SET(1)
	Return
}

;******************************************************************************
;  ���t�̓���
;******************************************************************************
; Excel�ł̓�������O����
If WinNotActive("ahk_class EXCEL.EXE")
{
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
}

;******************************************************************************
;  Excel�݂̂ŗL���ȃL�[
;******************************************************************************
If WinActive("ahk_class EXCEL.EXE")
{
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
}


;******************************************************************************
;  PowerPoint�݂̂ŗL���ȃL�[
;******************************************************************************
If WinActive("ahk_class POWERPNT.EXE")
{
	; Ctrl�{1�i�t���L�[�j �Ő}�`�̏���
	^1::
	{
		Send "{AppsKey}o{Enter}"
		Return
	}	
}

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
If WinNotActive("ahk_class Code.EXE")
{
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
}

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
;���ϊ�+m���X�N���[���_�E��
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





