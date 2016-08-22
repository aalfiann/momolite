# momolite

Apa itu momo library lite version?

momo library adalah framework minimalis dalam lingkungan .Net untuk membuat aplikasi Anda tipis, mudah dalam perawatan dan cepat dalam pengembangan. momolite adalah versi kecil dari momo library yang bebas digunakan secara gratis.

=================================================================================

Apa saja fitur dalam momo library lite version ini?

- Mendukung koneksi database untuk SQL Server, MySQL, SQLite, Firebird dan PostgreSQL.
- Dapat berjalan secara native atau portable tanpa harus instal konektor database.
- Tersedia fitur export dan import CSV, Excell, HTML, JSON, Text, TSV dan XML
- Mendukung enkripsi dan dekripsi SHA, AES, DES, 3DES dan MD5.
- Generate unique Serial Number
- Library dapat di gunakan dalam format 32/64bit 
- Otomatis generate error.log
- Dll.   

=================================================================================

Catatan untuk para developer dalam menggunakan momolite Library version 1.x:
1. Jika Anda menggunakan momo library lite version, pastikan Anda menambahkan referensi momolite.dll dan jangan lupa untuk menaruh: 

- mLib (Folder)
- momolite.dll
- momolite.dll.config 
- momolite.xml

di dalam root folder aplikasi Anda.

2. Jangan lupa untuk menaruh kode di bawah ini di dalam app.config aplikasi Anda, untuk memastikan bahwa library yang digunakan di momolite.dll akan berjalan di aplikasi Anda.

Contoh app.config untuk x86 sebagai berikut:
<?xml version="1.0" encoding="utf-8" ?>
	<configuration>
		<runtime>
			<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
		      <probing privatePath="mLib;mLib\32"/>
		    </assemblyBinding>
		</runtime>
	</configuration>

Contoh app.config untuk x64 sebagai berikut:
<?xml version="1.0" encoding="utf-8" ?>
	<configuration>
		<runtime>
			<assemblyBinding xmlns="urn:schemas-microsoft-com:asm.v1">
		      <probing privatePath="mLib;mLib\64"/>
		    </assemblyBinding>
		</runtime>
	</configuration>

3. Sekarang Anda dapat menggunakan momo library lite version dengan sempurna.

=================================================================================

Jika Anda menemukan bugs di dalam momo Library, maka hubungi saya di aalfiann@gmail.com.

Untuk cara penggunaan dan dokumentasi akan di posting di website http://javelinee.com

Terima kasih telah menggunakan momo Library lite version secara gratis.

==

M ABD AZIZ ALFIAN
Founder and Developer momo library
