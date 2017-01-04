$oMyError = ObjEvent("AutoIt.Error","MyErrFunc"); Install a custom error handler
Local $objDomain = ObjGet("WinNT://" & @ComputerName & "" )
Dim $filter[2] = ["user"]
$objDomain.Filter = $filter

$list = ""

For $aUser In $objDomain
    $list = $list & $aUSer.Name & "|" & $aUSer.Description & @CRLF
Next

MsgBox(0,"User Accounts",$list)