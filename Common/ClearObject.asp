<%
if isObject(RsUnit) then
    if RsUnit.state then
        RsUnit.close        
    end if
    Set RsUnit = Nothing
end if

if isObject(RsMember) then
    if RsMember.state then
        RsMember.close        
    end if
    Set RsMember = Nothing
end if


if isObject(RsMemberData) then
    if RsMemberData.state then
        RsMemberData.close
    end if
    Set RsMember = Nothing
end if

if isObject(RSGroup) then
    if RSGroup.state then
        RSGroup.close
    end if
    Set RSGroup = Nothing
end if

if isObject(RSSystem) then
    if RSSystem.state then
        RSSystem.close
    end if
    Set RSSystem = Nothing
end if

if isObject(RsUpd1) then
    if RsUpd1.state then
        RsUpd1.close        
    end if
    Set RsUpd1 = Nothing
end if

if isObject(RsUpd2) then
    if RsUpd2.state then
        RsUpd2.close        
    end if
    Set RsUpd2 = Nothing
end if

if isObject(RsUpd3) then
    if RsUpd3.state then
        RsUpd3.close        
    end if
    Set RsUpd3 = Nothing
end if

if isObject(RsTemp) then
    if RsTemp.state then
        RsTemp.close        
    end if
    Set RsTemp = Nothing
end if

if isObject(RsTemp2) then
    if RsTemp2.state then
        RsTemp2.close        
    end if
    Set RsTemp2 = Nothing
end if

if isObject(RsLoss) then
    if RsLoss.state then
        RsLoss.close        
    end if
    Set RsLoss = Nothing
end if

if isObject(RsChk) then
    if RsChk.state then
        RsChk.close
    end if
    Set RsChk = Nothing
end if

if isObject(RsMailHisotry) then
    if RsMailHisotry.state then
        RsMailHisotry.close
    end if
    Set RsMailHisotry = Nothing
end if

if isObject(RsBillMailHistory) then
    if RsBillMailHistory.state then
        RsBillMailHistory.close
    end if
    Set RsBillMailHistory = Nothing
end if


if isObject(RsReturnReson) then
    if RsReturnReson.state then
        RsReturnReson.close
    end if
    Set RsReturnReson = Nothing
end if

if isObject(RsLaw) then
    if RsLaw.state then
        RsLaw.close
    end if
    Set RsLaw = Nothing
end if

if isObject(RsStreet) then
    if RsStreet.state then
        RsStreet.close
    end if
    Set RsStreet = Nothing
end if

if isObject(Conn) then
    if Conn.state then    	  
        Conn.close
    end if
    Set Conn = Nothing
end if
%>