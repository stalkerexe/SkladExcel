Attribute VB_Name = "a_spr_all"

Public Sub load_mjj_all()
    On Error Resume Next

    ReDim mj(1 To 10, 1 To 1)
    
    mj(1, 1) = "Андреев Д.В."
    mj(2, 1) = "Кузнецов Д.В."
    mj(3, 1) = "Сидоров И.О."
    mj(4, 1) = "Брызгалин Р.Б."
    mj(5, 1) = "Тимошевич Ю.П."
    mj(6, 1) = "Беликов А.С."
    mj(7, 1) = "Забродин А.С."
    mj(8, 1) = "Черкашенинов А.Н."
    mj(9, 1) = "Анискин С.В."
    mj(10, 1) = "Ищенко Е.А."
    
End Sub


Public Sub load_zkz_all()
    On Error Resume Next

    ReDim zkz(1 To 7, 1 To 1)

    zkz(1, 1) = "ООО АГРОФИРМА «ПЯТИГОРЬЕ»"
    zkz(2, 1) = "ООО ГК «АЛЬФА-СПК-ДЖИТЕЙЧ»"
    zkz(3, 1) = "ООО Компания «Карат»"
    zkz(4, 1) = "ООО Контракт - Авто"
    zkz(5, 1) = "ООО НПФ «ТРЭКОЛ"""
    zkz(6, 1) = "ООО ТСК «ВОСТОК - СТРОЙМАРКЕТ»"
    zkz(7, 1) = "ООО Фирма «ТеплоЦель»"
End Sub


Public Sub load_doc_all()
    On Error Resume Next

    ReDim doc(1 To 4, 1 To 1)

    doc(1, 1) = "счет"
    doc(2, 1) = "счет-фактура"
    doc(3, 1) = "накладная"
    doc(4, 1) = "тов-трансп.наклад"

End Sub


