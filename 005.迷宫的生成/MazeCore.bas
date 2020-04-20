Attribute VB_Name = "MazeCore"
Public Type PosData
    X As Integer
    Y As Integer
End Type
Function RootFind(f() As PosData, X As Integer, Y As Integer) As PosData
    Dim ret As PosData
    If f(X, Y).X = X And f(X, Y).Y = Y Then ret = f(X, Y): GoTo last
    ret = RootFind(f, f(X, Y).X, f(X, Y).Y)
    f(X, Y) = ret
last:
    RootFind = ret
End Function
Function BuildMaze(W As Integer, H As Integer, Optional RoadWidth As Integer = 1, Optional Seed As Long = 404233) As Integer()
    Dim Map() As PosData
    Dim X As Integer, Y As Integer
    '初始化地图
    W = Int(W / 2) * 2: H = Int(H / 2) * 2
    ReDim Map(W, H)
    For X = 0 To W
        For Y = 0 To H
            Map(X, Y).X = X: Map(X, Y).Y = Y
            '墙壁
            If (X Mod 2 = 1) Or (Y Mod 2 = 1) Then Map(X, Y).X = -1
        Next
    Next
    Randomize Seed
    Dim M1 As PosData, M2 As PosData
    Dim T1 As PosData, T2 As PosData
    Dim Step As Boolean
    M1.X = -1
    
    Do While (M1.X <> M2.X Or M1.Y <> M2.Y)
        M1 = RootFind(Map, 0, 0): M2 = RootFind(Map, W, H)
        Step = Not Step
        X = Int(Rnd * (W + 2) / 2) * 2 + IIf(Step, 1, 0): Y = Int(Rnd * (H + 2) / 2) * 2 + IIf(Step, 0, 1)
        If X > W Then X = W
        If Y > H Then Y = H
        If Not ((X Mod 2 = 1) And (Y Mod 2 = 1)) Then
            '还是墙么
            If Map(X, Y).X = -1 Then
                If X Mod 2 = 1 Then '竖墙
                    '不会回路吧？
                    T1 = RootFind(Map, X - 1, Y): T2 = RootFind(Map, X + 1, Y)
                    If T1.X <> T2.X Or T1.Y <> T2.Y Then
                        Map(T1.X, T1.Y).X = T2.X: Map(T1.X, T1.Y).Y = T2.Y
                        Map(X, Y).X = 0
                    End If
                End If
                If Y Mod 2 = 1 Then '横墙
                    '不会回路吧？
                    T1 = RootFind(Map, X, Y - 1): T2 = RootFind(Map, X, Y + 1)
                    If T1.X <> T2.X Or T1.Y <> T2.Y Then
                        Map(T1.X, T1.Y).X = T2.X: Map(T1.X, T1.Y).Y = T2.Y
                        Map(X, Y).X = 0
                    End If
                End If
            End If
        End If
        'DoEvents
    Loop
    
    '格式转换
    Dim Maze() As Integer, RW As Integer
    RW = RoadWidth
    ReDim Maze(W * RW + 2, H * RW + 2)
    For X = 0 To W * RW + 2
        For Y = 0 To H * RW + 2
            If X = 0 Or Y = 0 Or X = W * RW + 2 Or Y = H * RW + 2 Then
                Maze(X, Y) = 1
            Else
                If Map((X - 1) / RW, (Y - 1) / RW).X = -1 Then
                    Maze(X, Y) = 1
                Else
                    Maze(X, Y) = 0
                End If
            End If
        Next
    Next
    
    BuildMaze = Maze
End Function
