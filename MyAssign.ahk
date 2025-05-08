#Requires AutoHotkey v2.0

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
; 2025/01/24 ひらがな/カタカナキーでもIMEをONにする
; 2025/03/07 AutoHotKey v2.0に対応
; 2025/03/10 特定のexeのみ有効な設定が効いていなかったのを修正、しかしexcelを除外する設定が常に効いてしまう問題あり
; 2025/04/11 まれに半角英数になって戻せなくなる問題を修正、誤字修正
; 2025/04/14 Win＋変換、Win＋カタカナでも半角英数になってしまう問題を修正
;            ただし現在、Shift＋カタカナで全角カタカナになってしまう問題が回避できていない
; 2025/05/08 無変換＋；（日付入力）のエクセル除外が効いていなかったのを修正
;            Shift＋Ctrl＋P（パスワード生成）のVSCode除外が効いていなかったのを修正
;--------------------------------------------------------
; キーアサイン一覧 
; 無変換+v　		バージョン表示
; 無変換		IME を OFF
; 変換			IME を ON
; カタカナ		IME を ON
; Ctrl+;		日付入力（yyyy/mm/dd形式）※エクセルを除く
; Ctrl+Shift+;		日付入力（yyyymmdd形式）※エクセルを除く
; Ctrl+Shift+V	書式なしペースト　※エクセルのみ
; Ctrl+1（フルキー）	図形の書式　※PowerPointのみ
; 
; 左ALT+h		左カーソル移動 ※シフト併用で選択
; 無変換+h		左カーソル移動 ※シフト併用で選択
; 
; 左ALT+j		下カーソル移動 ※シフト併用で選択
; 無変換+j		下カーソル移動 ※シフト併用で選択
; 
; 左ALT+k		上カーソル移動 ※シフト併用で選択
; 無変換+k		上カーソル移動 ※シフト併用で選択
; 
; 左ALT+l		右カーソル移動 ※シフト併用で選択
; 無変換+l		右カーソル移動 ※シフト併用で選択
; 
; 左ALT+u		PageDown
; 無変換+u		PageDown
; 
; 左ALT+i		PageUp
; 無変換+i		PageUp
; 
; 左ALT+y		Home
; 無変換+y		Home
; 
; 左ALT+o		End
; 無変換+o		End
; 
; Shift+Ctrl+p		パスワード生成 ※VSCodeを除く
; 
; 無変換+t		TAB文字をペースト
; 
; 無変換+m		スクロールダウン
; 無変換+,		スクロールアップ
; 無変換+n		左スクロール
; 無変換+.		右スクロール
;--------------------------------------------------------

#Include IMEv2.ahk

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
{
	Return
}

;;******************************************************************************
;;  テスト用
;;******************************************************************************
;vkF2::
;{
;	MsgBox("test")
;	return	
;}


;******************************************************************************
;  バージョン表示
;******************************************************************************
;無変換+v→バージョン表示
vk1D & v::
{
	MsgBox ("MyAssign last update 2025/05/08")
	return
}

;******************************************************************************
;  IMEのOn/Off
;******************************************************************************
; 無変換 で IME を OFF
vk1D::
 {
	IME_SET(0)
	Return
}
; 変換、カタカナ で IME を ON
vk1C::
vkF2::
{
	IME_SET(1)
	Return
}
; Ctrl＋カタカナ、Win＋カタカナ、Shift＋カタカナ、Win＋変換を無効化
^vkF2::
#vkF2::
+vkF2::
#vk1C::
{
	Return
}


;******************************************************************************
;  日付の入力
;******************************************************************************
; Excelでの動作を除外する
;#HotIf not WinActive("ahk_exe EXCEL.EXE") ; Excel以外がアクティブなとき
#HotIf WinGetProcessName("A") != "EXCEL.EXE"
	;Ctrl+;で日付入力(yyyy/mm/dd形式)
	^vkBB::
	{
		A_Clipboard:=A_YYYY "/" A_MM "/" A_DD
		Send "^v"
		Return
	}
	;Ctrl+Shift+;で日付入力(yyyymmdd形式)
	^+vkBB::
	{
		A_Clipboard:=A_YYYY A_MM A_DD
		Send "^v"
		Return
	}
#HotIf

;******************************************************************************
;  Excelのみで有効なキー
;******************************************************************************
#HotIf WinActive("ahk_exe EXCEL.EXE") ; Excelがアクティブなとき
	; F1ヘルプを無効化
	F1::
	{
		Return
	}	
	; Ctrl＋Shift＋V で書式なしペースト
	^+v::
	{
		Send "{AppsKey}v"
		Return
	}
#HotIf


;******************************************************************************
;  PowerPointのみで有効なキー
;******************************************************************************
#HotIf WinActive("ahk_exe POWERPNT.EXE")
	; Ctrl＋1（フルキー） で図形の書式
	^1::
	{
		Send "{AppsKey}o{Enter}"
		Return
	}	
#HotIf

;******************************************************************************
;  viのカーソル移動マップ
;******************************************************************************

;左ALT+h→左
<!h::
{	if GetKeyState("shift", "P"){
		Send "+{Left}"
	}else{
		Send "{Left}"
	}
	return
}
;左ALT+j→下
<!j::
{
	if GetKeyState("shift", "P"){
		Send "+{Down}"
	}else{
		Send "{Down}"
	}
	return
}
;左ALT+k→うえ
<!k::
{
	If GetKeyState("shift", "P"){
		Send "+{Up}"
	}else{
		Send "{Up}"
	}
	return
}
;左ALT+l→右
<!l::
{
	If GetKeyState("shift", "P"){
		Send "+{Right}"
	}else{
		Send "{Right}"
	}
	return
}
;無変換+h→左
vk1D & h::
{
	if GetKeyState("shift", "P"){
		Send "+{Left}"
	}else{
		Send "{Left}"
	}
	return
}
;無変換+j→下
vk1D & j::
{
	if GetKeyState("shift", "P"){
		Send "+{Down}"
	}else{
		Send "{Down}"
	}
	return
}
;無変換+k→うえ
vk1D & k::
{
	If GetKeyState("shift", "P"){
		Send "+{Up}"
	}else{
		Send "{Up}"
	}
	return
}
;無変換+l→右
vk1D & l::
{
	If GetKeyState("shift", "P"){
		Send "+{Right}"
	}else{
		Send "{Right}"
	}
	return
}
; left alt+yuioでPageDown/PageUp/Home/End
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

; 無変換+yuioでPageDown/PageUp/Home/End
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
; パスワード自動生成
;******************************************************************************
; VS Code での動作を除外する
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
;  TABの入力（Tabキーで次のコントロールに移動するのではなく、Tab文字をペーストする）
;******************************************************************************
;無変換+t
vk1D & t::
{
	A_Clipboard:=A_Tab
	Send "^v"
	return
}

;******************************************************************************
;  OneNoteのみの設定
;******************************************************************************
#HotIf WinActive("ahk_class Framework::CFrame")
    <!k::      DllCall("keybd_event", "int", 0x26, "int", 0, "int", 1, "int", 0) ;Up
    vk1D & k:: DllCall("keybd_event", "int", 0x26, "int", 0, "int", 1, "int", 0) ;Up
    <!j::      DllCall("keybd_event", "int", 0x28, "int", 0, "int", 1, "int", 0) ;Down
    vk1D & j:: DllCall("keybd_event", "int", 0x28, "int", 0, "int", 1, "int", 0) ;Down
#HotIf

;******************************************************************************
;  マウスホイール エミュレーション
;******************************************************************************
;無変換+m→スクロールダウン
vk1D & m::
{
	MouseClick "WheelDown"
	return
}
;無変換+,→スクロールダウン
vk1D & ,::
{
	MouseClick "WheelUp"
	return
}
;無変換+n→左スクロール
vk1D & n::
{
	MouseClick "WheelLeft"
	return
}
;無変換+.→右スクロール
vk1D & .::
{
	MouseClick "WheelRight"
	return
}
