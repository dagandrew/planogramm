#a::PickCat()

PickCat(){
 ;MouseClick("Left", 180, 84)
 ;Sleep(300)
 MouseClick("Left", 807, 67)
 Sleep(300)
 MouseClick("Left", 35, 378)
 Sleep(300)
 Send("^c")
}

!1::OpenWindow(1)
!2::OpenWindow(2)
!3::OpenWindow(3)
!4::OpenWindow(4)
!5::OpenWindow(5)

OpenWindow(number){
	Send("!{о}")
	sleep(300)
	Send(number)
}


;! o -> 1 or 2 открытые окна
;! с -> п -> enter планограммы
;! 3 Shift Tab -> Enter -> right -> down -> enter Шаблон то/но
