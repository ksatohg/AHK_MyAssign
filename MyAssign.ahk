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
;--------------------------------------------------------

; AutoHotkey: v1.1.24.05
; Author:     karakaram   http://www.karakaram.com/alt-ime-on-off
#Include IME.ahk

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
	Return


;******************************************************************************
;  �o�[�W�����\��
;******************************************************************************
;���ϊ�+v���o�[�W�����\��
vk1D & v::
MsgBox ,
(
MyAssign
last update 2024/01/23
)
return


;******************************************************************************
;  IME��On/Off
;******************************************************************************
; ���ϊ� �� IME �� OFF
vk1D::
    IME_SET(0)
    Return

; �ϊ� �� IME �� ON
vk1C::
    IME_SET(1)
    Return


;******************************************************************************
;  ���t�̓���
;******************************************************************************
; Excel�ł̓�������O����
#IfWinNotActive ahk_exe EXCEL.EXE
	;Ctrl+;�œ��t����(yyyy/mm/dd�`��)
	^vkBB::
		Clipboard=%A_YYYY%/%A_MM%/%A_DD%
		Send, ^v
		Return

	;Ctrl+Shift+;�œ��t����(yyyy/mm/dd�`��)
	^+vkBB::
		Clipboard=%A_YYYY%%A_MM%%A_DD%
		Send, ^v
		Return
#IfWinNotActive


;******************************************************************************
;  Excel�݂̂ŗL���ȃL�[
;******************************************************************************
#IfWinActive ahk_exe EXCEL.EXE
	; F1�w���v�𖳌���
	F1::
		Send {F2}
		Return
	; Ctrl�{Shift�{V �ŏ����Ȃ��y�[�X�g
	^+v::
		Send {AppsKey}v
		Return
#IfWinActive


;******************************************************************************
;  PowerPoint�݂̂ŗL���ȃL�[
;******************************************************************************
#IfWinActive ahk_exe POWERPNT.EXE
	; Ctrl�{1�i�t���L�[�j �Ő}�`�̏���
	^1::
		Send {AppsKey}o
		Return
#IfWinActive


;******************************************************************************
;  vi�̃J�[�\���ړ��}�b�v
;******************************************************************************

;��ALT+h����
<!h::
	if GetKeyState("shift", "P"){
		Send, +{Left}
	}else{
		Send,{Left}
	}
	return
;��ALT+j����
<!j::
	if GetKeyState("shift", "P"){
		Send, +{Down}
	}else{
		Send,{Down}
	}
	return
;��ALT+k������
<!k::
	If GetKeyState("shift", "P"){
		Send, +{Up}
	}else{
		Send,{Up}
	}
	return
;��ALT+l���E
<!l::
	If GetKeyState("shift", "P"){
		Send, +{Right}
	}else{
		Send,{Right}
	}
	return

;���ϊ�+h����
vk1D & h::
	if GetKeyState("shift", "P"){
		Send, +{Left}
	}else{
		Send,{Left}
	}
	return
;���ϊ�+j����
vk1D & j::
	if GetKeyState("shift", "P"){
		Send, +{Down}
	}else{
		Send,{Down}
	}
	return
;���ϊ�+k������
vk1D & k::
	If GetKeyState("shift", "P"){
		Send, +{Up}
	}else{
		Send,{Up}
	}
	return
;���ϊ�+l���E
vk1D & l::
	If GetKeyState("shift", "P"){
		Send, +{Right}
	}else{
		Send,{Right}
	}
	return

; left alt+yuio��PageDown/PageUp/Home/End
<!u::
	Send,{PgDn}
	return
<!i::
	Send,{PgUp}
	return
<!y::
	Send,{Home}
	return
<!o::
	Send,{End}
	return

; ���ϊ�+yuio��PageDown/PageUp/Home/End
vk1D & u::
	Send,{PgDn}
	return
vk1D & i::
	Send,{PgUp}
	return
vk1D & y::
	Send,{Home}
	return
vk1D & o::
	Send,{End}
	return

;******************************************************************************
; �p�X���[�h��������
;******************************************************************************
; VS Code �ł̓�������O����
#IfWinNotActive ahk_exe Code.EXE
	^+p::
		Number = 23456789
		Lowercase = abcdefghjkmnpqrstuvwxyz
		Uppercase = ABCDEFGHJKMNPQRSTUVWXYZ
		Mark = !#$`%@?+-*;=
		
		Password =
		Loop,8
		{
			Random,Start, 1, 4
			If Start = 1
			{
				Type = Number
			}
			If Start = 2
			{
				Type = Lowercase
			}
			If Start = 3
			{
				Type = Uppercase
			}
			If Start = 4
			{
				Type = Mark
			}
			
			StringLen, VLeng,  %Type%
			Random,Start, 1, %VLeng%
			StringMid, VCharacter, %Type%, Start,1,
			Password = %Password%%VCharacter%
		}
		Clipboard=%Password%
		Send, ^v
		return
#IfWinNotActive

;******************************************************************************
;  TAB�̓��̓o�[�W�����\��
;******************************************************************************
;���ϊ�+t
vk1D & t::
	Clipboard={0x09}
	Send,^v
	return


;******************************************************************************
;  OneNote�݂̂̐ݒ�
;******************************************************************************
#IfWinActive, ahk_class Framework::CFrame
    <!k::     dllcall("keybd_event", int, 0x26, int, 0, int, 1, int, 0) ;Up
    vk1D & k::dllcall("keybd_event", int, 0x26, int, 0, int, 1, int, 0) ;Up
    <!j::     dllcall("keybd_event", int, 0x28, int, 0, int, 1, int, 0) ;Down
    vk1D & j::dllcall("keybd_event", int, 0x28, int, 0, int, 1, int, 0) ;Down
#IfWinActive

;******************************************************************************
;  �}�E�X�z�C�[�� �G�~�����[�V����
;******************************************************************************
;���ϊ�+m���X�N���[���_�E��
vk1D & m::
	MouseClick, WheelDown
	return
;���ϊ�+m���X�N���[���_�E��
vk1D & ,::
	MouseClick, WheelUp
	return
;���ϊ�+n�����X�N���[��
vk1D & n::
	MouseClick, WheelLeft
	return
;���ϊ�+.���E�X�N���[��
vk1D & .::
	MouseClick, WheelRight
	return






