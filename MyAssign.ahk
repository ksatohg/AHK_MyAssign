;--------------------------------------------------------
; MyAssign made with AutoHotKey
;--------------------------------------------------------
; update history
;
; 2022/02/04 バックアップからWin10用に修正、IMEのon/offを削除
;            AHKのバージョンアップに伴い、vk1Dsc07B → vk1D のようにキー文字列を変更
; 2022/02/07 alt+k が一部の環境で正しく動作しなかったため、コメントを修正
;            無変換＋vでバージョンを表示するようにした
; 2022/06/28 Excel上の動作を改善
; 2022/07/04 IMEのOn/Offを復活
; 2022/07/06 Excel上の動作を改善
; 2022/07/07 Excelで書式なしペーストを追加
; 2022/07/07 ExcelでF1もF2にする
; 2023/03/09 SHIFT + CTRL + P（パスワード生成）をVSCode上で無効化（コマンドパレット表示と重複するため）
; 2023/03/28 テキストボックス内にTAB文字をペーストしようとしたがうまくいかない
; 2023/07/25 OneNoteで上下キーが反応しない問題に対応
;            無変換＋nm,.でマウスホイールエミュレーションを追加
; 2024/01/23 PowerPointで、図形の書式に対応
;--------------------------------------------------------

; AutoHotkey: v1.1.24.05
; Author:     karakaram   http://www.karakaram.com/alt-ime-on-off
#Include IME.ahk

; 主要なキーを HotKey に設定し、何もせずパススルーする
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
;*~vk1D:: ;無変換
;*~vk1C:: ;変換
	Return


;******************************************************************************
;  バージョン表示
;******************************************************************************
;無変換+v→バージョン表示
vk1D & v::
MsgBox ,
(
MyAssign
last update 2024/01/23
)
return


;******************************************************************************
;  IMEのOn/Off
;******************************************************************************
; 無変換 で IME を OFF
vk1D::
    IME_SET(0)
    Return

; 変換 で IME を ON
vk1C::
    IME_SET(1)
    Return


;******************************************************************************
;  日付の入力
;******************************************************************************
; Excelでの動作を除外する
#IfWinNotActive ahk_exe EXCEL.EXE
	;Ctrl+;で日付入力(yyyy/mm/dd形式)
	^vkBB::
		Clipboard=%A_YYYY%/%A_MM%/%A_DD%
		Send, ^v
		Return

	;Ctrl+Shift+;で日付入力(yyyy/mm/dd形式)
	^+vkBB::
		Clipboard=%A_YYYY%%A_MM%%A_DD%
		Send, ^v
		Return
#IfWinNotActive


;******************************************************************************
;  Excelのみで有効なキー
;******************************************************************************
#IfWinActive ahk_exe EXCEL.EXE
	; F1ヘルプを無効化
	F1::
		Send {F2}
		Return
	; Ctrl＋Shift＋V で書式なしペースト
	^+v::
		Send {AppsKey}v
		Return
#IfWinActive


;******************************************************************************
;  PowerPointのみで有効なキー
;******************************************************************************
#IfWinActive ahk_exe POWERPNT.EXE
	; Ctrl＋1（フルキー） で図形の書式
	^1::
		Send {AppsKey}o
		Return
#IfWinActive


;******************************************************************************
;  viのカーソル移動マップ
;******************************************************************************

;左ALT+h→左
<!h::
	if GetKeyState("shift", "P"){
		Send, +{Left}
	}else{
		Send,{Left}
	}
	return
;左ALT+j→下
<!j::
	if GetKeyState("shift", "P"){
		Send, +{Down}
	}else{
		Send,{Down}
	}
	return
;左ALT+k→うえ
<!k::
	If GetKeyState("shift", "P"){
		Send, +{Up}
	}else{
		Send,{Up}
	}
	return
;左ALT+l→右
<!l::
	If GetKeyState("shift", "P"){
		Send, +{Right}
	}else{
		Send,{Right}
	}
	return

;無変換+h→左
vk1D & h::
	if GetKeyState("shift", "P"){
		Send, +{Left}
	}else{
		Send,{Left}
	}
	return
;無変換+j→下
vk1D & j::
	if GetKeyState("shift", "P"){
		Send, +{Down}
	}else{
		Send,{Down}
	}
	return
;無変換+k→うえ
vk1D & k::
	If GetKeyState("shift", "P"){
		Send, +{Up}
	}else{
		Send,{Up}
	}
	return
;無変換+l→右
vk1D & l::
	If GetKeyState("shift", "P"){
		Send, +{Right}
	}else{
		Send,{Right}
	}
	return

; left alt+yuioでPageDown/PageUp/Home/End
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

; 無変換+yuioでPageDown/PageUp/Home/End
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
; パスワード自動生成
;******************************************************************************
; VS Code での動作を除外する
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
;  TABの入力バージョン表示
;******************************************************************************
;無変換+t
vk1D & t::
	Clipboard={0x09}
	Send,^v
	return


;******************************************************************************
;  OneNoteのみの設定
;******************************************************************************
#IfWinActive, ahk_class Framework::CFrame
    <!k::     dllcall("keybd_event", int, 0x26, int, 0, int, 1, int, 0) ;Up
    vk1D & k::dllcall("keybd_event", int, 0x26, int, 0, int, 1, int, 0) ;Up
    <!j::     dllcall("keybd_event", int, 0x28, int, 0, int, 1, int, 0) ;Down
    vk1D & j::dllcall("keybd_event", int, 0x28, int, 0, int, 1, int, 0) ;Down
#IfWinActive

;******************************************************************************
;  マウスホイール エミュレーション
;******************************************************************************
;無変換+m→スクロールダウン
vk1D & m::
	MouseClick, WheelDown
	return
;無変換+m→スクロールダウン
vk1D & ,::
	MouseClick, WheelUp
	return
;無変換+n→左スクロール
vk1D & n::
	MouseClick, WheelLeft
	return
;無変換+.→右スクロール
vk1D & .::
	MouseClick, WheelRight
	return






