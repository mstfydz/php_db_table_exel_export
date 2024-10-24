<?php
include '../settings.php';
require '../vendor/SimpleXLSXGen.php'; 
use Shuchkin\SimpleXLSXGen;

try {
    // Eğer Excel'e aktarım isteği varsa
    if (isset($_GET['start_date']) && isset($_GET['end_date'])) {
        // Tarih aralığını al
        $start_date = $_GET['start_date'];
        $end_date = $_GET['end_date'];

        // Tarihleri saat bilgisi ile birleştir
        $start_date = date('Y-m-d H:i:s', strtotime($start_date . ' 00:00:00'));
        $end_date = date('Y-m-d H:i:s', strtotime($end_date . ' 23:59:59'));

        // SQL sorgusunu hazırla (sadece istediğiniz sütunları seçin)
        $stmt = $db->prepare("SELECT * FROM talepler WHERE kayit_tarih BETWEEN :start_date AND :end_date ORDER BY kayit_tarih ASC");
        $stmt->bindParam(':start_date', $start_date);
        $stmt->bindParam(':end_date', $end_date);
        $stmt->execute();
        $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

        // Verileri başlık satırı ile birleştir
        $data = [];
        if (!empty($rows)) {
            // Başlık satırını ekle
            $data[] = [
                'Ad Soyad', 
                'TC Kimlik No', 
                'E-posta Adresi', 
                'Tel No', 
                'Unvan', 
                'Sözleşme', 
                'Yeni başvuru / Yenileme', 
                'Kurulum Tarihi', 
                'Personel Adı', 
                'Sertifika Bedeli', 
                'Fatura No', 
                'Ödeme Şekli', 
                'Satış Noktası', 
                'Sipariş No', 
                'Nakit Takip'
            ];

            // Verileri satır satır ekle
            foreach ($rows as $row) {
                // Başvuru şekline göre değer belirle
                switch ($row['basvurusekli']) {
                    case 1:
                        $basvurusekli = "Yeni Başvuru";
                        break;
                    case 2:
                        $basvurusekli = "Yenileme";
                        break;
                    default:
                        $basvurusekli = "Tanımsız";
                        break;
                }

                // Ad ve soyadı birleştir
                $fullName = trim($row['adi'] . ' ' . $row['soyadi']);
                // Verileri diziye ekle
                $data[] = [
                       '<left><style font-size="10">' . $fullName . '</style></left>',
                       '<center><style font-size="10" font-family="Calibri">' . $row['tckn'] . '</style></center>', // Sadece bu hücre Calibri olacak
                       '<center><style font-size="11">'.$row['email'].'</style></center>',
                       '<center><style font-size="10">'.$row['telefon'].'</style></center>',
                       '',  // Unvan
                       '<center><style font-size="10">'.$row['sure'].'</style></center>',
                       '<center><style font-size="10">'.$basvurusekli.'</style></center>',
                       '<center><style font-size="10">'. date(format: 'Y-m-d', strtotime($row['kayit_tarih'])) .'</style></center>',
                       '<left><style font-size="10">'. $row['personel'] .'</style></left>',
                       '',  // Sertifika Bedeli
                       '',  // Fatura No
                       '',  // Ödeme Şekli
                       '<left><style font-size="10">'. $row['personel'] .'</style></left>',
                       '',  // Sipariş No
                       ''
                ];	

            }
        }

        // Yeni bir XLSX dosyası oluşturun
        $xlsx = SimpleXLSXGen::fromArray($data);
            $xlsx->setDefaultFont('Arial');
            $xlsx->setDefaultFontSize(10);
            // Excel dosyasını indirin
            $xlsx->download('talepler.xlsx');
        exit; // Export işlemi tamamlandığında çıkış yap
    }
} catch (PDOException $e) {
    echo "Veritabanı bağlantı hatası: " . $e->getMessage();
}
?>