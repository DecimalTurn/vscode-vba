' SYNTAX TEST "source.vba" "metadata"

VERSION 1.0 CLASS
BEGIN
' <----- keyword.metadata.vba
  MultiUse = -1  'True
            '^^  constant.numeric.decimal
                '^^^^^  comment.line.quote
END

Sub test()

    MsgBox "test"
    End
   '^^^ keyword.control.vba

End Sub
