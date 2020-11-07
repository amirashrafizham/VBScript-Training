'*******************************************************
'Script Name: Captain Adventure.vbs
'Author: Amir Ashraf
'Created: 7/11/2020
'Description: This script prompsts the user to answer a number of questions_
'and then uses the answers to create a comical action adventure story
'*******************************************************
Option Explicit

Dim name
Dim place
Dim object
Dim friend
Dim dessert
Dim story
Dim title

title = "Captain Adventure"

MsgBox      "Welcome to the story of........."&_
            vbCrLf&"CCC          A"&_
            vbCrLf&"C             AAA"&_
            vbCrLf&"CCCaptain  A   Adventure gets his super powers",vbOkOnly,title

name = InputBox ("What is your name")
place = InputBox ("Name a place you would like to visit")
object = InputBox ("Name a strange object")
friend = InputBox ("Type the name of a close friend")
dessert = InputBox ("Type the name of your favourite dessert")


story = "Once upon a time"&_
        vbCrLf&" "&_
        vbCrLf&name&" went on a vacation in the far away land of "&place&". A local tour guide suggested cave exploration. While in the cave "&name&" accidentally became separated from the rest of the "&_
        "group and stumbled into a  part of the cave never visited before. It was completely dark. Suddenly a powerful light began to glow."&name&" saw that it came from a mysterious "&object&" located "&_
        "in the far corner of the cave room. "&name&" picked it up and a flash of light occured and "&name&" was instantly transported to a far away world. There in front of him was "&friend&", the ancient"&_
        "God of the Legendary Cave People. "&friend&" explained to "&name&" that destiny had selected him to become Captain Adventure! He was then returned to Earth and told to purchase a"&dessert& "and "&_
        "travel the countryside looking for people in need of help. To activvate the superpowers bestowed by "&friend&" all that "&name&" had to do was pick up the "&object&" and say Picklethree 3 times in a row"

MsgBox story,vbOkOnly,title