
'CodeFile1 نوع الكلاس
Namespace FormTools

    Public Class ResizeControls

        Dim RatioTable As New Hashtable

        Private WindowHeight As Single
        Private WindowWidth As Single
        Private HeightRatio As Single
        Private WidthRatio As Single

        Private _Container As New Control

        Public Property Container() As Control
            Get
                Return _Container
            End Get
            Set(ByVal Ctrl As Control)
                _Container = Ctrl
                FullRatioTable()
            End Set

        End Property

        Private Structure SizeRatio
            Dim TopRatio As Single
            Dim LeftRatio As Single
            Dim HeightRatio As Single
            Dim WidthRatio As Single
        End Structure

        Private Sub FullRatioTable()
            WindowHeight = _Container.Height
            WindowWidth = _Container.Width
            RatioTable = New Hashtable
            AddChildrenToTable(_Container)
        End Sub

        Private Sub AddChildrenToTable(ByRef ChildContainer As Control)
            Dim R As New SizeRatio
            For Each C As Control In ChildContainer.Controls
                With C
                    R.TopRatio = CSng(.Top / WindowHeight)
                    R.LeftRatio = CSng(.Left / WindowWidth)
                    R.HeightRatio = CSng(.Height / WindowHeight)
                    R.WidthRatio = CSng(.Width / WindowWidth)
                    RatioTable(.Name) = R
                    If .HasChildren Then
                        AddChildrenToTable(C)
                    End If
                End With
            Next
        End Sub

        Public Sub ResizeControls()

            HeightRatio = CSng(_Container.Height / WindowHeight)
            WidthRatio = CSng(_Container.Width / WindowWidth)

            WindowHeight = _Container.Height
            WindowWidth = _Container.Width

            ResizeChildren(_Container)

        End Sub

        Private Sub ResizeChildren(ByRef ChildContainer As Control)
            Dim R As New SizeRatio
            For Each C As Control In ChildContainer.Controls
                With C
                    R = CType(RatioTable(.Name), FormTools.ResizeControls.SizeRatio)
                    .Top = CInt(WindowHeight * R.TopRatio)
                    .Left = CInt(WindowWidth * R.LeftRatio)
                    .Height = CInt(WindowHeight * R.HeightRatio)
                    .Width = CInt(WindowWidth * R.WidthRatio)
                    If .HasChildren Then
                        ResizeChildren(C)
                    End If

                    Select Case True
                        Case TypeOf C Is ListBox
                            Dim L As New ListBox
                            L = CType(C, ListBox)
                            L.IntegralHeight = False

                        Case TypeOf C Is ListView
                            ResizeColumnsL(C, WidthRatio)

                        Case TypeOf C Is DataGridView
                            ResizeColumnsD(C, WidthRatio)

                        Case TypeOf C Is DataGrid
                            ResizeColumnsDg(C, WidthRatio)
                    End Select


                    ResizeControlFont(C, WidthRatio, HeightRatio)
                End With
            Next
        End Sub

        Private Sub ResizeControlFont(ByRef Ctrl As Control, ByVal RatioW As Single, ByVal RatioH As Single)

            Dim FSize As Single = Ctrl.Font.Size
            Dim FStyle As FontStyle = Ctrl.Font.Style
            Dim FNome As String = Ctrl.Font.Name
            Dim NewSize As Single = FSize

            If TypeOf Ctrl Is DataGridView Then
                Dim D As DataGridView

                D = CType(Ctrl, DataGridView)
                D.DefaultCellStyle.Font = New Font(D.Font.FontFamily, CSng(FSize * Math.Sqrt(RatioW * RatioH)), FontStyle.Regular)

            End If

            NewSize = CSng(FSize * Math.Sqrt(RatioW * RatioH))
            Dim NFont As New Font(FNome, CSng(NewSize), FStyle)
            Ctrl.Font = NFont

        End Sub

        Private Sub ResizeColumnsL(ByRef Ctrl As Control, ByVal RatioW As Single)

            Dim C As ColumnHeader
            For Each C In CType(Ctrl, ListView).Columns
                C.Width = CInt(C.Width * RatioW)
            Next

        End Sub

        Private Sub ResizeColumnsD(ByRef Ctrl As Control, ByVal RatioW As Single)

            Dim C As DataGridViewColumn
            For Each C In CType(Ctrl, DataGridView).Columns
                C.Width = CInt(C.Width * RatioW)
            Next

        End Sub

        Private Sub ResizeColumnsDg(ByRef Ctrl As Control, ByVal RatioW As Single)

            Dim C As DataGridColumnStyle
            For Each C In CType(Ctrl, DataGrid).TableStyles(0).GridColumnStyles
                C.Width = CInt(C.Width * RatioW)
            Next

        End Sub

    End Class

End Namespace