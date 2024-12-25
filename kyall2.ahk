SendMode("Input")  ; Recommended for faster key sending

windows := WinGetList() ; Get a list of window handles
result := "" ; Initialize result string

for id in windows
{
    wt := WinGetTitle(id) ; Get the title of each window
    result .= wt . "`n" ; Append the title to the result string
}

MsgBox(result) ; Display the result
wordWindows := WinGetList("Word")
; Define hotkeys to capture key presses and send them to all open Word windows
~a::SendToAll("a",wordWindows)
~b::SendToAll("b",wordWindows)
~c::SendToAll("c",wordWindows)
~d::SendToAll("d",wordWindows)
~e::SendToAll("e",wordWindows)
~f::SendToAll("f",wordWindows)
~g::SendToAll("g",wordWindows)
~h::SendToAll("h",wordWindows)
~i::SendToAll("i",wordWindows)
~j::SendToAll("j",wordWindows)
~k::SendToAll("k",wordWindows)
~l::SendToAll("l",wordWindows)
~m::SendToAll("m",wordWindows)
~n::SendToAll("n",wordWindows)
~o::SendToAll("o",wordWindows)
~p::SendToAll("p",wordWindows)
~q::SendToAll("q",wordWindows)
~r::SendToAll("r",wordWindows)
~s::SendToAll("s",wordWindows)
~t::SendToAll("t",wordWindows)
~u::SendToAll("u",wordWindows)
~v::SendToAll("v",wordWindows)
~w::SendToAll("w",wordWindows)
~x::SendToAll("x",wordWindows)
~y::SendToAll("y",wordWindows)
~z::SendToAll("z",wordWindows)
~1::SendToAll("1",wordWindows)
~2::SendToAll("2",wordWindows)
~3::SendToAll("3",wordWindows)
~4::SendToAll("4",wordWindows)
~5::SendToAll("5",wordWindows)
~6::SendToAll("6",wordWindows)
~7::SendToAll("7",wordWindows)
~8::SendToAll("8",wordWindows)
~9::SendToAll("9",wordWindows)
~0::SendToAll("0",wordWindows)
~.::SendToAll(".",wordWindows)
~,::SendToAll(",",wordWindows)
~?::SendToAll("?",wordWindows)
~!::SendToAll("{!}",wordWindows)
~#::SendToAll("{#}",wordWindows)
~+::SendToAll("{+}",wordWindows)
~^::SendToAll("{^}",wordWindows)
~-::SendToAll("-",wordWindows)
~{::SendToAll("{{}",wordWindows)
~}::SendToAll("{}}",wordWindows)
~(::SendToAll("(",wordWindows)
~)::SendToAll(")",wordWindows)
~[::SendToAll("[",wordWindows)
~]::SendToAll("]",wordWindows)
~Space::SendToAll(" ",wordWindows)  ; Space
~Enter::SendToAll("{Enter}",wordWindows)  ; Enter
~Backspace::SendToAll("{Backspace}",wordWindows)  ; Backspace
~Tab::SendToAll("{Tab}",wordWindows)  ; Tab

~Delete::SendToAll("{Delete}",wordWindows)  ; Delete
~Up::SendToAll("{Up}",wordWindows)  ; Arrow up
~Down::SendToAll("{Down}",wordWindows)  ; Arrow down
~Left::SendToAll("{Left}",wordWindows)  ; Arrow left
~Right::SendToAll("{Right}",wordWindows)  ; Arrow right

~Esc::ExitApp()
; Function to send the key to all open Word windows
SendToAll(Key, wordWindows) {
	; ControlFocus(wordWindows[1])
	; ControlSend(Key,"_WwG1","Document3 - Word")
	ControlSend(Key,"_WwG1","WARFD2P_Main_12192024 - Word")
}