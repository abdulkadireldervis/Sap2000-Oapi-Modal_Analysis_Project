import os
import sys
import comtypes.client
import csv

# SAP2000 uygulamasına bağlanma
AttachToInstance = False
SpecifyPath = False
ProgramPath = 'C:\\Program Files\\Computers and Structures\\SAP2000 25\\SAP2000.exe'
APIPath = input('Proje dosyalarının kaydedileceği klasör yolunu girin: ')
if not os.path.exists(APIPath):
    os.makedirs(APIPath)
ModelPath = APIPath + os.sep + 'Modal_Analysis_Model.sdb'

# API yardımcı nesnesi oluşturma
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

if AttachToInstance:
    try:
        mySapObject = helper.GetObject('CSI.SAP2000.API.SapObject')
    except (OSError, comtypes.COMError):
        print('SAP2000 çalışmıyor veya bağlanılamadı.')
        sys.exit(-1)
else:
    if SpecifyPath:
        mySapObject = helper.CreateObject(ProgramPath)
    else:
        mySapObject = helper.CreateObjectProgID('CSI.SAP2000.API.SapObject')
    mySapObject.ApplicationStart()

# SAP2000 model nesnesini oluşturma
SapModel = mySapObject.SapModel
SapModel.InitializeNewModel()
SapModel.File.NewBlank()

# Birimleri metre-kilonewton olarak ayarlama
kN_m_C = 6
SapModel.SetPresentUnits(kN_m_C)

# C30/37 Beton Malzeme Tanımlama
MATERIAL_NAME = 'C30/37'
MATERIAL_CONCRETE = 2  # 2: Beton
SapModel.PropMaterial.SetMaterial(MATERIAL_NAME, MATERIAL_CONCRETE)
SapModel.PropMaterial.SetMPIsotropic(MATERIAL_NAME, 3e7, 0.2, 0.0000055)
SapModel.PropMaterial.SetWeightAndMass(MATERIAL_NAME, 25, 0)

# Kesit Tanımlamaları
# Kolon Kesiti: 60 cm x 30 cm
SapModel.PropFrame.SetRectangle('KOLON', MATERIAL_NAME, 0.6, 0.3)

# Kiriş Kesiti: 30 cm x 60 cm
SapModel.PropFrame.SetRectangle('KIRIS', MATERIAL_NAME, 0.6, 0.3)

# 2D çerçeve modelinin geometri tanımları
FrameNames = []

# Düğüm noktaları ve çerçeve elemanları tanımlama
# Kolonlar
[FrameName1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 3, '', 'KOLON', '1', 'Global')
[FrameName2, ret] = SapModel.FrameObj.AddByCoord(5, 0, 0, 5, 0, 3, '', 'KOLON', '2', 'Global')
[FrameName3, ret] = SapModel.FrameObj.AddByCoord(0, 0, 3, 0, 0, 6, '', 'KOLON', '3', 'Global')
[FrameName4, ret] = SapModel.FrameObj.AddByCoord(5, 0, 3, 5, 0, 6, '', 'KOLON', '4', 'Global')

# Kirişler
[FrameName5, ret] = SapModel.FrameObj.AddByCoord(0, 0, 3, 5, 0, 3, '', 'KIRIS', '5', 'Global')
[FrameName6, ret] = SapModel.FrameObj.AddByCoord(0, 0, 6, 5, 0, 6, '', 'KIRIS', '6', 'Global')

# Sabit mesnetler tanımlama (alt düğüm noktaları)
SabitMesnet = [True, True, True, False, False, False]
SapModel.PointObj.SetRestraint('1', SabitMesnet)
SapModel.PointObj.SetRestraint('3', SabitMesnet)

# Modeli kaydetme
SapModel.File.Save(ModelPath)

print('2D çerçeve modeli başarıyla oluşturuldu, C30/37 beton malzeme ve kesitler tanımlandı ve model kaydedildi!')


# Sabit mesnetler tanımlama (alt düğüm noktaları)
SabitMesnet = [True, True, True, False, False, False]
SapModel.PointObj.SetRestraint('1', SabitMesnet)
SapModel.PointObj.SetRestraint('5', SabitMesnet)

# Modal analiz yükleme durumunu tanımlama
SapModel.LoadCases.ModalEigen.SetCase('MODAL')

SapModel.LoadCases.ModalEigen.SetNumberModes('MODAL', 10, 2)  # 10 mod şekli hesaplanacak

# Modal analizi çalıştırma
ret = SapModel.Analyze.RunAnalysis()
if ret != 0:
    print(f'Modal analiz başarısız oldu! Hata kodu: {ret}')
else:
    print('Modal analiz başarıyla tamamlandı!')

# Modal periyot ve frekansları alma
NumberResults, LoadCase, StepType, StepNum, Period, Frequency, CircFreq, EigenValue, ret = SapModel.Results.ModalPeriod(
    0, [], [], [], [], [], [], []
)

# Sonuçları terminalde gösterme
if NumberResults > 0:
    for i in range(NumberResults):
        print(f'Mod {i+1}: Periyot={Period[i]} sn, Frekans={Frequency[i]} Hz, Dairesel Frekans={CircFreq[i]} rad/sn')
else:
    print('Modal analiz sonuçları bulunamadı.')
