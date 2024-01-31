package com.data;

import java.util.*;
import java.io.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class Siswa {
    String nama;
    String nisn;
    int no, nilaiIndonesia, nilaiMatematika, nilaiInggris, nilaiProduktif;

    Siswa(int no, String nama, String nisn, int nilaiIndonesia, int nilaiMatematika, int nilaiInggris,
            int nilaiProduktif) {
        this.no = no;
        this.nama = nama;
        this.nisn = nisn;
        this.nilaiIndonesia = nilaiIndonesia;
        this.nilaiMatematika = nilaiMatematika;
        this.nilaiInggris = nilaiInggris;
        this.nilaiProduktif = nilaiProduktif;
    }

    int getNilaiTotal() {
        return nilaiIndonesia + nilaiMatematika + nilaiInggris + nilaiProduktif;
    }
}

public class Main {
    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("Enter the path of the Excel file:");
        String excelFilePath = scanner.nextLine();

        List<Siswa> siswaList = readExcelFile(excelFilePath);

        // Cetak semua ranking siswa
        cetakHasil(siswaList);

        // Minta pengguna memasukkan NISN
        System.out.println("Enter NISN of the student to search:");
        String nisn = scanner.nextLine();

        // Cari siswa berdasarkan NISN dan cetak detailnya
        Siswa siswa = cariSiswa(siswaList, nisn);
        if (siswa != null) {
            int ranking = siswaList.indexOf(siswa) + 1;
            // Print details to console or any other desired output
            printDetailsToConsole(siswa, ranking);
        } else {
            System.out.println("Student with NISN " + nisn + " not found.");
        }

        scanner.close();
    }

    static List<Siswa> readExcelFile(String excelFilePath) throws IOException {
        List<Siswa> siswaList = new ArrayList<>();
        FileInputStream fis = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            // Skip header row
            int currentRow = row.getRowNum();
            if (currentRow != 0 && currentRow != 1) {
                int no = getNumericValueFromCell(row.getCell(0));
                String nama = getStringValueFromCell(row.getCell(1));
                String nisn = getStringValueFromCell(row.getCell(2));
                int nilaiIndonesia = getNumericValueFromCell(row.getCell(4));
                int nilaiMatematika = getNumericValueFromCell(row.getCell(5));
                int nilaiInggris = getNumericValueFromCell(row.getCell(6));
                int nilaiProduktif = getNumericValueFromCell(row.getCell(7));

                siswaList.add(new Siswa(no, nama, nisn, nilaiIndonesia, nilaiMatematika, nilaiInggris, nilaiProduktif));
            }
        }
        workbook.close();
        fis.close();
        return siswaList;
    }

    static String getStringValueFromCell(Cell cell) {
        return cell != null ? cell.getStringCellValue() : "";
    }

    static int getNumericValueFromCell(Cell cell) {
        return cell != null ? (int) cell.getNumericCellValue() : 0;
    }

    static void printDetailsToConsole(Siswa siswa, int ranking) {
        System.out.println("Nama: " + siswa.nama);
        System.out.println("NISN: " + siswa.nisn);
        System.out.println("Nilai Indonesia: " + siswa.nilaiIndonesia);
        System.out.println("Nilai Matematika: " + siswa.nilaiMatematika);
        System.out.println("Nilai Inggris: " + siswa.nilaiInggris);
        System.out.println("Nilai Produktif: " + siswa.nilaiProduktif);
        System.out.println("------------------------");
        System.out.println("Nilai Total: " + siswa.getNilaiTotal());
        System.out.println("Ranking: " + ranking);
        System.out.println("------------------------");
    }

    static void cetakHasil(List<Siswa> siswaList) throws IOException {
        // Urutkan siswa berdasarkan nilai total
        Collections.sort(siswaList, (a, b) -> b.getNilaiTotal() - a.getNilaiTotal());

        // Buat PrintWriter untuk menulis ke file
        PrintWriter writer = new PrintWriter(new FileWriter("hasil.txt"));

        // Cetak hasil ke file
        for (int i = 0; i < siswaList.size(); i++) {
            Siswa siswa = siswaList.get(i);
            writer.println("Nama: " + siswa.nama);
            writer.println("NISN: " + siswa.nisn);
            writer.println("Nilai Indonesia: " + siswa.nilaiIndonesia);
            writer.println("Nilai Matematika: " + siswa.nilaiMatematika);
            writer.println("Nilai Inggris: " + siswa.nilaiInggris);
            writer.println("Nilai Produktif: " + siswa.nilaiProduktif);
            writer.println("------------------------");
            writer.println("Nilai Total: " + siswa.getNilaiTotal());
            writer.println("Ranking: " + (i + 1));
            writer.println("------------------------");
        }

        // Tutup PrintWriter
        writer.close();
    }

    static Siswa cariSiswa(List<Siswa> siswaList, String nisn) {
        for (Siswa siswa : siswaList) {
            if (siswa.nisn.equals(nisn)) {
                return siswa;
            }
        }
        return null;
    }
}
