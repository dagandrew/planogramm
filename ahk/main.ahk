#Requires AutoHotkey v2.0

F1::^c
F2::^a
F3::^v
F4::^x
F2 & F1::CopyAll()
F2 & F3::ReplaceAll()

` & 1::{
	Send("^{s}")
	Reload
}

`::Home
CapsLock::End
` & 2::^+!I ;screenshot scissors (desktop shortcut)

CopyAll(){
	Send("^{a}")
	Send("^{c}")
}

ReplaceAll(){
	Send("^{a}")
	Send("^{v}")
}

#include excel.ahk
#include planogramm.ahk
