gci ENV:\ComputerName
gci -path c:\ -File -recurse -include *.pst -erroraction 'silentlycontinue'|select-object fullname,lastwritetime,length | fl fullname,lastwritetime,length
gci -path d:\ -File -recurse -include *.pst -erroraction 'silentlycontinue'|select-object fullname,lastwritetime,length | fl fullname,lastwritetime,length
gci -path e:\ -File -recurse -include *.pst -erroraction 'silentlycontinue'|select-object fullname,lastwritetime,length | fl fullname,lastwritetime,length
