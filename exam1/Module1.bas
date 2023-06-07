Attribute VB_Name = "Module1"
Private monster As New CHamster
Private myMonster As New CMyHamster

Public Sub monster_eat()
    monster.eat
End Sub

Public Sub monster_play()
    monster.play
End Sub

Public Sub monster_work()
    monster.work
End Sub

Public Sub SetMoster()
    Set monster = myMonster
End Sub
