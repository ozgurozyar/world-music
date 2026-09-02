/* ==========================================================================
   Ozgur World — talep formu veri dilimi
   ELLE YAZILMAZ. Excel'deki Web_Export sekmesinden üretilir.

   TAZELEME (kütüphane her güncellendiğinde):
   1) Excel'de Web_Export sekmesini aç
   2) D3'teki tarihi al, aşağıdaki OW_TARIH satırına yaz
   3) G6 hücresine tıkla → Ctrl+Shift+Aşağı Ok → Ctrl+C
   4) Aşağıdaki ROWS bloğunun içine yapıştır (iki işaret satırının ARASINA)
   5) I6 hücresine tıkla → Ctrl+Shift+Aşağı Ok → Ctrl+C
   6) SYN bloğunun içine yapıştır
   7) Satır sayısını denetle: Web_Export!F3 kaç yazıyorsa ROWS bloğunda o kadar
      satır olmalı. Tutmuyorsa yapıştırma eksik — tekrar dene.

   Alan sırası (boru ile ayrılmış): Bölge|Ana Tür|Mood|İşlev|Detay Tarz
   Parça adı ve URI BİLİNÇLİ OLARAK YOK — bu dosya herkese açık duruyor.
   ========================================================================== */

var OW_TARIH = "2026-08-31";

var OW_ROWS = `
/* --- G SÜTUNU BURAYA --- */
Iranian World|Blues|Melankoli|Dans; Gece; Yolculuk|Fars Füzyon; Blues; Klasik İran Müziği
Western Europe; Anatolia; Levant; Fusion|Caz; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Klarnet + Ud + Perküsyon; Türk/Arap Repertuvarı; Almanya
Levant; Mediterranean Europe; Fusion|Caz; Tel/Enstrümantal|Neşe|Dans; Odak; Yolculuk|Ud + Akordeon + Tuba; Arap Makamı – Avrupa Cazı
Levant; Western Europe; Fusion|Caz|Neşe|Dans; Odak; Yolculuk|Oud Cazı; Arap Makamı; Avrupa Kamara Cazı
North Africa; Levant; Iberia; Fusion|Elektronik|Melankoli|Gece|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Caribbean|Vokal/Şarkı|Romantik|Gece|Son Cubano; Bolero; Kuba Geleneği
Western Europe; Latin America; Fusion|Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Elektro-Tango; Downtempo
North America|Elektronik|Neşe|Dans; Odak; Spor|Downtempo/Dub; Bossa-Reggae Ödünçlemeleri; Washington DC
North America|Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Downtempo/Dub; Bossa-Reggae Ödünçlemeleri; Washington DC
North Africa|Elektronik|Romantik|Dans; Yolculuk; Spor|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Balkans|Nefesli/Bando|Neşe|Dans|Makedon Roman Nefesli Orkestrası; Koçani
North Africa|Elektronik|Neşe|Dans; Spor|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Levant; North Africa; Iberia; Fusion|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Erken Müzik; Oud-Viyol; Akdeniz Diyaloğu
North America|Psikedelik/Rock|Neşe|Dans; Spor|Gypsy Punk; Balkan Rock; Göçmen Kabare
Iberia|Gelenek|Neşe|Odak; Yolculuk|Flamenko; Utrera Ekolü; Kuba Tres Gitarı
West Africa|Blues|Dinginlik|Dans; Odak; Spor|Mali Blues; Afro-Pop; Gitar Düeti
Anatolia; North America|Elektronik; Gelenek|Karanlık|Dans; Yolculuk; Spor|Ney + Elektronik; Sufi Esini; Arkın Allen
North Africa|Elektronik|Neşe|Dans; Yolculuk; Spor|Raï Rock; Elektronik Chaâbi; Punk
North Africa|Elektronik|Neşe|Dans; Yolculuk; Spor|Raï Rock; Elektronik Chaâbi; Punk
North Africa|Elektronik|Neşe|Dans|Raï Rock; Elektronik Chaâbi; Punk
Levant|Vokal/Şarkı|Melankoli|Gece|Arap Şarkı Geleneği; Rahbani Besteleri; Beyrut
Balkans|Nefesli/Bando|Neşe|Dans; Odak; Spor|Balkan Brass; Roman Nefesli Orkestrası
West Africa; Fusion|Vokal/Şarkı|Romantik|Yolculuk|Mande Vokali; Morna; Afrika Diva Düeti
Caribbean|Funk/Soul|Neşe|Dans|Afro-Küba Füzyon; Latin Hip-Hop; New York Groove
North America|Caz|Romantik|Dans; Gece|Retro Pop Orkestrası; Çok Dilli Kabare; Latin Caz
North America|Caz|Romantik|Dans; Yolculuk; Spor|Retro Pop Orkestrası; Çok Dilli Kabare; Latin Caz
Iberia|Gelenek; Hip-Hop/Rap|Neşe|Dans; Spor|Flamenko Füzyon; Rap/Ragga; Barselona
Iberia|Gelenek|Neşe|Dans; Spor|Endulüs-Mağrip Füzyonu; Erken Müzik; Flamenko
Latin America|Vokal/Şarkı|Neşe|Dans; Spor|Meksika Halk Müziği; Ranchera; Tex-Mex
Western Europe; Levant; Fusion|Blues|Neşe|Yolculuk|Arap-Blues Füzyonu; Oud-Slide Gitar; İki Gelenek
Caucasus|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Azeri Mugam; Aşık Geleneği; Kafkas Vokali
Western Europe; South Asia; Fusion|Psikedelik/Rock|Epik|Dans; Yolculuk|Tamil-Blues Füzyon; Sufi Şarkısı; Alternatif Rock
Latin America|Vokal/Şarkı|Romantik|Dans; Yolculuk|Slowcore Chanson; Meksika Bolero; Alternatif Folk
Western Europe|Pop/Şarkı-Yazarı|Neşe|Spor|Çağdaş Chanson; Manouche/Sokak Tavır
Western Europe|Pop/Şarkı-Yazarı|Neşe|Yolculuk|Çağdaş Chanson; Manouche/Sokak Tavır
Anatolia|Caz|Epik|Dans; Odak; Yolculuk|Anadolu Cazı; Perküsyon; Serbest Doğaçlama
Anatolia|Caz|Epik|Odak; Spor|Anadolu Cazı; Perküsyon; Serbest Doğaçlama
Anatolia|Caz|Neşe|Yolculuk|Anadolu Cazı; Perküsyon; Serbest Doğaçlama
East & Central Asia; Northern Europe; Fusion|Caz|Neşe|Dans; Odak; Yolculuk; Spor|Caz Vokal; Kamara Caz; Kore Etkisi
Western Europe|Pop/Şarkı-Yazarı; Nefesli/Bando|Melankoli|Dans; Yolculuk; Spor|Patchanka; Alternatif Chanson; 80'ler-90'lar
Balkans; North America; Fusion|Psikedelik/Rock; Nefesli/Bando|Karanlık|Spor|Balkan Punk/Metal; Seattle; Göç Anlatısı
Balkans; North America; Fusion|Psikedelik/Rock; Nefesli/Bando|Neşe|Dans; Spor|Balkan Punk/Metal; Seattle; Göç Anlatısı
Balkans; North America; Fusion|Psikedelik/Rock; Nefesli/Bando|Karanlık|Spor|Balkan Punk/Metal; Seattle; Göç Anlatısı
Latin America; North America|Psikedelik/Rock|Melankoli|Gece|Meksika Rock'ı; Mariachi + Gitar; Film Müziği
North Africa|Vokal/Şarkı|Melankoli|Gece|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
North Africa|Pop/Şarkı-Yazarı|Romantik|Dans|Fransız-Cezayirli Pop; Chanson + Mağrip Renkleri
Western Europe|Klasik/Modern Klasik|Melankoli|Odak; Gece|Minimalist Piyano; Film Müziği; Breton Folk
Iberia|Gelenek|Neşe|Dans; Spor|Rumba Catalana; Roman Geleneği; Kurucu Figur
South Asia|Vokal/Şarkı|Neşe|Odak|Playback Şarkıcılığı; Bollywood; Hint Film Müziği
Western Europe|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Yolculuk; Spor|French Pop; Synth-Pop; Electronica
Levant|Psikedelik/Rock|Karanlık|Dans; Spor|Arap Indie Rock; Alternatif; Politik Pop
South Asia; Western Europe; Fusion|Hip-Hop/Rap; Elektronik|Neşe|Dans; Odak; Spor|Bhangra + Hip-Hop; Punjabi Diaspora
Levant; Iberia; Fusion|Gelenek; Vokal/Şarkı|Romantik|Dans|Sefarad/Ladino; Flamenko; İspanyol Şarkısı; Akdeniz Füzyonu
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Electro Swing; Nu Jazz; Sample Bazlı Prodüksiyon
North Africa|Elektronik|Melankoli|Gece|Arap Alternatif; Protest Şarkı; Elektronik
North Africa|Elektronik|Karanlık|Gece|Arap Alternatif; Protest Şarkı; Elektronik
Levant|Psikedelik/Rock; Vokal/Şarkı|Neşe|Yolculuk; Spor|Akdeniz Surf Rock + Vokal
Levant|Caz|Epik|Odak; Gece|Contrabas Caz; Akdeniz Cazı; World Jazz
Caucasus|Caz; Klasik/Modern Klasik|Dinginlik|Odak; Gece; Yolculuk|Jazz Piano; Armenian Modalism; Contemporary Composition
North America|Caz; Vokal/Şarkı|Romantik|Gece|Caz Vokali; Rus/Yidiş Repertuvar; Toronto
Levant|Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Film Müziği; Nadine Labaki Filmleri; Beyrut
Northern Europe|Elektronik|Melankoli|Gece|Downtempo; İskandinav Elektronik; Vokal Konuğu
Iranian World|Vokal/Şarkı; Gelenek|Romantik|Dans; Spor|Fars Klasik Vokali (Âvâz); Kamanche-Tanbur
Levant|Tel/Enstrümantal; Klasik/Modern Klasik|Dinginlik|Odak; Yolculuk|Ud; Mahmud Derviş Şiirleri; Beyrut
Western Europe|Caz|Neşe|Dans|Balkan Swing; Caz Orkestrası; Çok Dilli Vokal
Western Europe; Levant; Fusion|Blues; Gelenek|Neşe|Dans|'Arabicana'; Ud + Blues Gitarı; Utrecht
Western Europe|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Piyano; Minimalist Repertuvar; Brilliant Classics
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|French Pop; Chanson; Variété Française
Iberia|Caz|Romantik|Dans|Flamenko-Caz; Kopla; Afro-İber Füzyon
Western Europe|Pop/Şarkı-Yazarı|Epik|Dans; Spor|Fransız Pop; İspanyolca Repertuvar; Protest Şarkı
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Balkan Brass; Elektronik; Sokak Bandosu
Mediterranean Europe; Balkans|Nefesli/Bando|Neşe|Spor|Balkan Nefesli; Güney İtalya Bandosu; Puglia
Western Europe|Gelenek|Melankoli|Dans|Klezmer; Roman Müziği; Yidiş Füzyon
Western Europe|Gelenek|Neşe|Odak|Klezmer; Roman Müziği; Yidiş Füzyon
Balkans|Nefesli/Bando|Neşe|Dans; Odak|Balkan Brass; Roman Nefesli Orkestrası
Balkans|Nefesli/Bando; Blues|Neşe|Dans|Balkan Brass; Blues Cover; Guest Vocal
Balkans|Nefesli/Bando; Gelenek|Neşe|Yolculuk; Spor|Roman Köy Orkestrası; Balkan Brass; Çoček
Balkans|Nefesli/Bando|Neşe|Dans; Spor|Balkan Brass; Roman Nefesli Orkestrası
Levant|Psikedelik/Rock|Neşe|Dans; Spor|Fransız-Arap Pop; Rock; Chanson Füzyon
Western Europe; Levant; Fusion|Blues; Gelenek|Melankoli|Dans|'Arabicana'; Ud + Blues Gitarı; Utrecht
Anatolia|Pop/Şarkı-Yazarı; Caz|Neşe|Dans; Spor|Piyano-Vokal; Türk Ezgileri + Caz; Amsterdam
Western Europe|Elektronik|Neşe|Dans|Balkan Elektronik; Bucovina Club; Yunan Vokal Konuğu
Anatolia|Elektronik; Psikedelik/Rock|Dinginlik|Odak; Gece; Yolculuk|Anadolu Elektroniği; Saz + Prodüksiyon; İstanbul
Mediterranean Europe|Klasik/Modern Klasik|Melankoli|Yolculuk|Keman/Besteci; Napoli; Film Müziği
Arabian Peninsula; Levant|Blues|Neşe|Dans; Yolculuk|Yemen Folk; Funk; Blues-Ortadoğu Füzyonu
Anatolia|Gelenek|Neşe|Dans; Yolculuk; Spor|Sufi Müziği; Ney/Zurna; Ortadoğu Füzyonu
Western Europe; West Africa; Fusion|Elektronik; Gelenek|Epik|Dans; Yolculuk; Spor|Kelt + Batı Afrika + Elektronik; Real World
West Africa|Blues; Tel/Enstrümantal|Melankoli|Yolculuk|Çöl Blues; Songhai Gitarı; Niafunké
Anatolia|Klasik/Modern Klasik|Dinginlik|Dans; Odak; Yolculuk; Spor|Piyano; Klasik-Anadolu Füzyon; Beste
Anatolia|Klasik/Modern Klasik|Neşe|Odak|Piyano; Klasik-Anadolu Füzyon; Beste
Levant; Iranian World; South Asia; Fusion|Gelenek; Tel/Enstrümantal|Romantik|Odak; Yolculuk; Spor|Ud/Keman + Fars Ney + Sarod + Darbuka; Doğu-Doğu Füzyonu; Tel Aviv Canlı Kaydı
Caucasus|Caz; Klasik/Modern Klasik|Dinginlik|Odak; Yolculuk|Mugam-Caz Piyanosu; Azeri Vokal; Bakü
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Fransız Chanson; Kolektif Kayıt; Yardım Projesi
North Africa|Elektronik|Karanlık|Dans; Yolculuk|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Northern Europe|Caz; Gelenek|Neşe|Odak|Gypsy Jazz; Zanzibar-Arap Repertuvar; Bouzouki-Keman
West Africa|Vokal/Şarkı|Romantik|Yolculuk|Morna; Coladeira; Cape Verde Geleneği
Levant|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Arabic Pop; Dance Pop; Lebanese Song; Electronic Remix
Balkans|Nefesli/Bando|Neşe|Dans; Odak; Spor|Balkan Brass; Roman Nefesli Orkestrası
Western Europe; Latin America; Fusion|Elektronik|Karanlık|Dans; Yolculuk; Spor|Elektro-Tango; Downtempo
Balkans|Nefesli/Bando|Melankoli|Odak; Gece|Balkan Brass; Roman Nefesli Orkestrası; Kolo
Caucasus; Anatolia|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Ermeni Repertuvarı; Perküsyon ve Vokal; Armenian Navy Band
Iberia|Gelenek; Vokal/Şarkı|Neşe|Dans|Flamenko Cantesi; Morente Ailesi; Granada
North Africa|Elektronik|Dinginlik|Odak; Gece; Yolculuk|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Iberia|Klasik/Modern Klasik|Dinginlik|Gece|Erken Müzik; Arp ve Vokal; Katalan-İskandinav Füzyon
West Africa|Blues|Neşe|Dans; Spor|Mali Blues; Afro-Pop; Gitar Düeti
North America|Caz|Romantik|Gece|Retro Pop Orkestrası; Çok Dilli Kabare; Latin Caz
North America|Caz|Romantik|Yolculuk|Retro Pop Orkestrası; Çok Dilli Kabare; Latin Caz
Western Europe|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Dance Pop; Electro; Club; Latin Pop; DJ Production
Anatolia|Gelenek|Neşe|Dans; Gece|Ege-Balkan Repertuvar; Akordeon; Çok Dilli Halk Müziği
Western Europe|Klasik/Modern Klasik|Romantik|Gece; Yolculuk|Ortaçağ Müziği; Balkan-Sefarad Repertuvar; Erken Müzik
North America|Psikedelik/Rock|Dinginlik|Odak; Gece|Freak Folk; Psikedelik Folk; İki Dilli Şarkı
Latin America|Gelenek|Epik|Odak|Nueva Canción; And Halk Müziği; Politik Şarkı
North America|Gelenek|Neşe|Dans|Çok Dilli Folk; Tango-Balkan Füzyon; Politik Şarkı
North America; Eastern Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Alt Tonlu Vokal; Doğu Avrupa Melankolisi; eski adı Chinawoman
Iranian World|Psikedelik/Rock|Melankoli|Dans; Gece; Yolculuk|Fars Alternatif Rock; Indie
Western Europe|Klasik/Modern Klasik|Epik|Odak; Gece|Piyano Üçlüsü; Oda Müziği; Çağdaş Repertuvar
Levant; West Africa; Fusion|Tel/Enstrümantal|Romantik|Gece|İsrail-Mali Füzyon; Piyano-Gitar Düeti
Balkans|Psikedelik/Rock|Neşe|Dans; Spor|Balkan Rock; Reggae Etkisi; Politik Şarkı
Iberia|Klasik/Modern Klasik|Melankoli|Gece|Minimalist Kamara; Film Müziği; Portekiz Yeni Klasik
Anatolia|Klasik/Modern Klasik|Neşe|Dans|Osmanlı Klasik Müziği; Caz Füzyonu; Kamara Topluluğu
Iranian World|Pop/Şarkı-Yazarı|Neşe|Dans|Fars Pop (düşük güven)
Balkans|Gelenek; Vokal/Şarkı|Romantik|Gece; Yolculuk|Batı Balkan Halk Müziği; Sevdalinka; Akustik; Zagreb
Eastern Europe|Klasik/Modern Klasik|Melankoli|Odak; Gece|Minimalist Piyano; Neo-Klasik; Kamara
North America|Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|İndie Folk-Rock; 10,000 Maniacs Çevresi
Mediterranean Europe|Vokal/Şarkı|Melankoli|Yolculuk|Sardinya Halk Müziği; Akdeniz Vokal Geleneği
Levant|Caz|Melankoli|Odak; Gece|Contrabas Caz; Akdeniz Cazı; World Jazz
Balkans|Reggae/Dub|Romantik|Dans; Spor|Yunan Ska/Reggae; Atina
Western Europe|Blues|Melankoli|Dans|Hollanda Bluesu; 'Bensonhurst Blues'; 70'ler
Caribbean|Vokal/Şarkı|Epik|Dans; Gece|Trova; Devrim Şarkıları; Guaracha
Levant|Klasik/Modern Klasik|Dinginlik|Odak; Gece; Yolculuk|Oud Üçlüsü; Arap Klasik Müziği
Northern Europe|Vokal/Şarkı; Pop/Şarkı-Yazarı|Dinginlik|Dans|Litvanya Vokali; İndie/Pop Hattı
Western Europe; Latin America; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Venezuelan Work Song; Latin House; Afro-Latin Percussion
Balkans|Pop/Şarkı-Yazarı|Romantik|Gece|Avrupa Popı; Aphrodite's Child; Yunan Vokali
Balkans|Elektronik|Neşe|Dans; Spor|Eski Yunan Kayıtlarının Yeniden Kurgusu; Hip-Hop/Swing; Atina
Anatolia|Gelenek|Melankoli|Gece|Çok Dilli Anadolu Repertuvarı; Halk Müziği Kolektifi
Western Europe|Caz|Neşe|Dans; Yolculuk|Balkan Swing; Caz Orkestrası; Çok Dilli Vokal
Balkans|Elektronik|Melankoli|Gece|Downtempo/Elektronik; Selanik
Iberia|Klasik/Modern Klasik; Elektronik|Neşe|Dans; Spor|Akordeon/Besteci; Retro-Sinematik; Bask Ülkesi
Eastern Europe|Vokal/Şarkı; Gelenek|Epik|Yolculuk|Slav Polifonisi; Kadın Vokal Topluluğu; Alan Kayıtları
Iranian World|Elektronik|Dinginlik|Yolculuk|Sufi Elektronik; Fars-Hint Vokal; Trans
Iranian World|Vokal/Şarkı; Elektronik|Neşe|Dans; Yolculuk|Fars/Ortadoğu Vokali; Santur; Niyaz & Vas
Mediterranean Europe|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans|Napoli Şarkısı; Akdeniz/Latin Repertuvarı; Tiyatro
North Africa|Vokal/Şarkı|Melankoli|Yolculuk|Sahrawi Repertuvarı; Sürgün/Direniş Şarkıları
Iranian World; North America|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Dans; Gece|Farsça İndie-Pop; Animasyon; Vancouver
Anatolia|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Piyanoda Türk Ezgileri
Western Europe|Funk/Soul; Pop/Şarkı-Yazarı|Karanlık|Dans; Spor|Noir Soul; Boğuk Vokal; Eindhoven
Anatolia|Gelenek|Melankoli|Odak; Gece; Yolculuk|Sufi Müziği; Ney/Zurna; Ortadoğu Füzyonu
Anatolia|Elektronik|Melankoli|Dans; Odak; Yolculuk; Spor|Elektronik; Anadolu Psychedelia; Organik House
North America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Gece|Amerikan Folk Canlanması; Protest Şarkı; 60'lar
Anatolia|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Özgün Müzik; Anatolian Folk; Protest Song; Film Music
West Africa|Gelenek|Dinginlik|Dans; Spor|Kora; Mande Geleneği; Griot Müziği
North Africa|Blues|Romantik|Dans; Gece|Gnawa Füzyon; Blues; Soul
Western Europe|Blues|Melankoli|Gece|Amerikana; Folk-Blues; Çok Uluslu Grup
Western Europe; Levant|Pop/Şarkı-Yazarı|Dinginlik|Odak; Gece|Fransız-İbrani Şanson; Fısıltılı Vokal; Paris
Western Europe|Pop/Şarkı-Yazarı|Romantik|Gece|Çağdaş Şanson; Bossa Etkisi; Paris
Latin America; Western Europe; Fusion|Vokal/Şarkı|Neşe|Gece|Kamara Folk; Çok Dilli Vokal Düeti
Western Europe|Funk/Soul; Pop/Şarkı-Yazarı|Neşe|Gece|Soul-Folk; Nijerya-Roman Kökeni
West Africa|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Bamana Repertuvarı; Ngoni + Elektrikli Gitar; Bağımsız Söz
Western Europe|Pop/Şarkı-Yazarı|Neşe|Gece|Post-Punk Kapaklarının Bossa Yorumu; Paris
Western Europe; North Africa; Fusion|Elektronik|Karanlık|Yolculuk|Elektronik; Arap Vokal; Trip-Hop
Anatolia|Psikedelik/Rock; Gelenek|Dinginlik|Odak; Gece|Kürt Folk-Rock; Kürtçe/Türkçe/Zazaca Repertuvar; İstanbul; Elektrik Gitar; Def
Mediterranean Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Dans; Yolculuk|İndie Folk-Rock; İngilizce Sözler; Urtovox Records
Latin America|Elektronik|Karanlık|Dans|Deneysel Latin Pop; Kolaj; Tango-Elektronik
Anatolia|Psikedelik/Rock; Pop/Şarkı-Yazarı|Neşe|Odak|Türkçe Rock; Türkü Yorumları; İstanbul
Anatolia|Caz; Vokal/Şarkı|Neşe|Gece|Caz-Soul Vokali; İstanbul
Anatolia; Levant; Fusion|Gelenek; Vokal/Şarkı|Romantik|Dans; Spor|İstanbul Sokak Sahnesi; Santur + Gitar + Vokal
Caucasus; Western Europe; Anatolia; Fusion|Hip-Hop/Rap|Neşe|Dans; Yolculuk; Spor|Ermeni Halk Müziği; Anadolu Füzyon; Hip-Hop
Balkans|Vokal/Şarkı|Neşe|Dans; Spor|Yunan Entekno/Laiko; Sahne Vokali; Atina
Latin America|Gelenek|Melankoli|Dans; Spor|Nueva Canción; And Halk Müziği; Politik Şarkı
Eastern Europe|Funk/Soul|Neşe|Spor|Instrumental Soul; Funk
North America|Psikedelik/Rock|Romantik|Dans; Odak|Çöl Rock; Tex-Mex/Mariachi Nefeslileri; Tucson
Western Europe; East & Central Asia; Fusion|Tel/Enstrümantal|Neşe|Dans; Spor|Gitar + Morin Khuur; Moğol İşbirliği; Film Müziği
Northern Europe|Pop/Şarkı-Yazarı|Melankoli|Odak; Yolculuk|İskandinav Şarkı-Yazarı; Orkestral Düzenleme
North America|Gelenek|Melankoli|Dans|Çok Dilli Folk; Tango-Balkan Füzyon; Politik Şarkı
Balkans|Elektronik|Neşe|Yolculuk|Eski Yunan Kayıtlarının Yeniden Kurgusu; Hip-Hop/Swing; Atina
Iberia|Klasik/Modern Klasik|Romantik|Odak; Gece; Yolculuk|Erken Müzik; Arp ve Vokal; Katalan-İskandinav Füzyon
Levant|Caz|Romantik|Dans|Caz Gitarı ve Vokal; New York
West Africa; Iberia; Fusion|Tel/Enstrümantal; Vokal/Şarkı|Melankoli|Gece|Kora-Çello Düeti; Griot Geleneği; İber Vokal; Katalan Vokali; Flamenko-Caz; Fado Etkisi; Minimal Vokal-Gitar; Katalan Şarkısı; Deneysel Prodüksiyon
Western Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Yolculuk; Spor|Singer-Songwriter; Acoustic Pop; Chamber Folk; Piano Ballad
Northern Europe|Pop/Şarkı-Yazarı; Klasik/Modern Klasik|Melankoli|Dans; Gece|Oda-Pop; Piyano/Arp; Berlin
Mediterranean Europe|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans; Spor|Dream Pop; Nu Soul; Dance Pop
Iranian World|Hip-Hop/Rap|Neşe|Dans; Spor|Persian Rap; Political Lyrics
North Africa|Caz|Dinginlik|Dans; Yolculuk|Hassani; Sahra Blues; Caz-Soul Füzyon
Levant|Psikedelik/Rock|Neşe|Dans|Fransız-Arap Pop; Rock; Chanson Füzyon
Levant|Vokal/Şarkı; Caz|Romantik|Gece|Suriye-Ermeni Vokali; Muvaşşah + Caz Düzenleme; Şam
Levant|Klasik/Modern Klasik|Melankoli|Odak; Gece; Yolculuk|Oud Üçlüsü; Arap Klasik Müziği
Balkans|Caz; Blues|Neşe|Dans; Yolculuk|Swing; Dixieland; Jazz; Rhythm and Blues
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Akordeon-Chanson; Paris; Mizahi Söz
Balkans|Gelenek; Vokal/Şarkı|Romantik|Gece|Makedon Geleneksel Müziği; Balkan Dünya Müziği; Caz Etkili Düzenlemeler
North America|Vokal/Şarkı; Gelenek|Melankoli|Dans; Yolculuk|Sefarad/Ladino Repertuvarı; New Mexico
Anatolia|Klasik/Modern Klasik|Dinginlik|Odak|Piyanoda Türk Ezgileri
Levant|Psikedelik/Rock|Neşe|Dans; Spor|Fransız-Arap Pop; Rock; Chanson Füzyon
Balkans|Pop/Şarkı-Yazarı|Romantik|Dans|Yunan Indie-Folk; İngilizce/Yunanca Söz
Anatolia|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Anadolu Makam Geleneği; Bağlama-Ud
Mediterranean Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Dans; Spor|İtalyan Popı; Geniş Vokal Aralığı; Cantautore Repertuvarı
Iranian World|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Soprano; Fars Klasik Repertuvarı; Diaspora
North Africa|Vokal/Şarkı|Melankoli|Dans; Yolculuk|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
North Africa|Vokal/Şarkı|Dinginlik|Dans|Kabil Şarkısı; Amazigh Folk; Akustik Gitar
Northern Europe|Pop/Şarkı-Yazarı|Epik|Spor|İskandinav Şarkı-Yazarı; Orkestral Düzenleme
Iberia|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Fado; Portekiz Vokal Geleneği
Balkans|Vokal/Şarkı; Gelenek|Melankoli|Gece|Yunan Halk/Laiko Vokali; Selanik
Anatolia|Elektronik; Caz|Melankoli|Yolculuk; Spor|Caz; Elektronik; Nublu Sahnesi
Levant; North Africa; Fusion|Reggae/Dub|Epik|Dans; Gece|Psychedelic Reggae; Ortadoğu Dub; Tuareg Konuğu
Northern Europe|Caz; Pop/Şarkı-Yazarı|Neşe|Dans|Nordik Caz Vokali; Şarkı-Yazarlığı
North Africa|Caz|Dinginlik|Odak; Gece; Yolculuk|Oud; Arap Klasik Müziği; Kamara Caz
Anatolia|Gelenek|Neşe|Dans; Odak; Yolculuk; Spor|Trakya Roman Klarneti; Düğün Müziği; Rumeli Repertuvar
Anatolia; Western Europe; Fusion|Reggae/Dub|Karanlık|Dans; Spor|Anadolu Psychedelia; Dub; Elektrikli Saz
Levant|Psikedelik/Rock|Neşe|Dans; Spor|Fransız-Arap Pop; Rock; Chanson Füzyon
Levant; Western Europe; Fusion|Caz; Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Arap Indie; Trompet Cazı; Canlı İşbirliği
Western Europe|Blues|Melankoli|Dans|Amerikana; Folk-Blues; Çok Uluslu Grup
North Africa|Blues|Neşe|Dans; Spor|Gnawa Füzyon; Blues; Soul
Southern Africa; Western Europe; Fusion|Caz|Melankoli|Odak; Yolculuk|Angola Tango; Afro-Tango; Piyano Üçlüsü
North America|Psikedelik/Rock|Epik|Dans|Balkan-Punk; Göçmen Rock; Kabare
Balkans|Klasik/Modern Klasik|Melankoli|Odak; Gece|Film Müziği; Yaylı Orkestra; Yunan Kamara
North America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Romantik|Dans; Gece|Şarkı-Yazarı Geleneği; Caz Vokali
Levant; Western Europe; Fusion|Klasik/Modern Klasik; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Arabic Oud; Taqsim; Samai; Jazz; Flamenco
Western Europe; West Africa|Pop/Şarkı-Yazarı; Funk/Soul|Melankoli|Gece|Afro-Soul/Pop; Berlin
Balkans|Elektronik|Karanlık|Dans; Spor|Yunan Retro Sample; Rebetiko-Hip-Hop; Laiko
Latin America|Elektronik; Vokal/Şarkı|Melankoli|Gece|Brezilya Deneysel Pop'u; Londra Yerleşik
Western Europe|Elektronik|Melankoli|Dans|Klezmer Füzyon; Balkan; Elektronik Folk
Levant; Western Europe; South Asia; Fusion|Psikedelik/Rock|Epik|Dans; Odak|Kavvali; Racastan Halk Müziği; Sufi-Rock Füzyon
North America|Elektronik; Caz|Neşe|Dans; Spor|Trip-Hop; Acid Jazz; Nu Jazz; Downtempo
Mediterranean Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Dans; Yolculuk|İndie Folk-Rock; İngilizce Sözler; Urtovox Records
Iranian World|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Fars Blues; Setar-Vokal; Deneysel İran Müziği
Balkans|Klasik/Modern Klasik|Melankoli|Yolculuk|Film Müziği; Angelopoulos İşbirliği; ECM
Latin America|Vokal/Şarkı|Epik|Dans; Odak; Gece|Nueva Canción; Şili Halk Şarkısı; Politik Repertuvar
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans; Spor|Bossa-Cover Pop; Post-Punk Yeniden Yorum; Kabare
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Caz; Trompet-Gitar Düeti; Kamara Caz
Western Europe|Caz|Melankoli|Dans; Odak|Ermeni Ezgileri; Caz Piyanosu; Şarkı Yazımı
Levant|Pop/Şarkı-Yazarı; Blues|Melankoli|Dans|Folk-Rock; İngilizce Söz; Tel Aviv
Levant|Elektronik|Melankoli|Yolculuk|Arap Trip-Hop; Beyrut Underground; Lo-Fi Elektronik
Western Europe|Elektronik|Romantik|Odak; Gece|Fransız Disko; Nu-Funk; Elektronik Pop
Western Europe; Mediterranean Europe|Elektronik; Vokal/Şarkı|Dinginlik|Odak; Yolculuk|Akdeniz Vokali + Elektronik Prodüksiyon
Levant|Caz|Neşe|Dans; Odak; Yolculuk|Mandolin; Ortadoğu Cazı; Kamara Füzyon
Western Europe|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Modern Klasik; Minimalist Piyano; Erken Ambient
Western Europe; South Asia; Fusion|Psikedelik/Rock|Melankoli|Dans|Tamil-Blues Füzyon; Sufi Şarkısı; Alternatif Rock
Mediterranean Europe|Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|Cantautore Geleneği; Cenova; Sardinya Esini
Mediterranean Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Cantautore Geleneği; Cenova; Sardinya Esini
North Africa|Caz|Dinginlik|Gece|Hassani; Sahra Blues; Caz-Soul Füzyon
North America|Gelenek|Neşe|Dans; Gece|Çok Dilli Folk; Tango-Balkan Füzyon; Politik Şarkı
Anatolia|Klasik/Modern Klasik|Neşe|Yolculuk|Anadolu Enstrümantal; Saha Kayıtları; Film Müziği
Anatolia|Klasik/Modern Klasik|Melankoli|Dans; Yolculuk; Spor|Anadolu Enstrümantal; Saha Kayıtları; Film Müziği
Arabian Peninsula; Levant; North America|Blues; Caz|Neşe|Dans; Spor|Yemen Folk; Funk; Blues-Ortadoğu Füzyonu; Mandolin; Ortadoğu Cazı; Kamara Füzyon; Kontrbas; Ortadoğu Ezgileri + Caz; New York
South Asia; Western Europe; Fusion; North Africa|Tel/Enstrümantal|Dinginlik|Dans; Gece; Yolculuk|Racastan Roman Müziği; Kalbeliya Dansı; Gitar-Oud Füzyon; Buzuk/Ud/Gitar; Roman-Hint-Mağrip Hattı
North America|Blues|Karanlık|Yolculuk|Güney Gotik Blues; Folk-Rock; Gospel Etkisi
Western Europe|Elektronik|Karanlık|Dans; Odak; Yolculuk|Electronic Remix; Global Bass
Iberia|Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Flamenko Gitarı; Enstrümantal; Londra Yerleşik
Western Europe; North Africa; Fusion|Elektronik|Karanlık|Yolculuk|Elektronik; Arap Vokal; Trip-Hop
Western Europe|Blues; Psikedelik/Rock|Neşe|Dans|Garage Blues; Lo-Fi; Thee Headcoatees Çevresi
Western Europe|Elektronik|Dinginlik|Yolculuk|Trip-Hop; Downtempo; 90'lar Britanya
Balkans|Vokal/Şarkı|Melankoli|Gece|Yunan Entekno/Laiko; Söz Merkezli Yorum
Western Europe|Caz; Elektronik|Melankoli|Odak; Yolculuk|İskandinav Cazı + Turntablism; Piyano-Theremin; Bristol Sahnesi
Anatolia|Gelenek; Vokal/Şarkı|Neşe|Spor|Rebetiko; Café Aman; Aegean Song; Film Soundtrack
North Africa|Vokal/Şarkı|Melankoli|Gece|Kabil Şarkısı; Amazigh Folk; Akustik Gitar
North America|Blues; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Blues + Klezmer + Tuba; New York
Anatolia|Psikedelik/Rock|Neşe|Odak; Spor|Anadolu Psychedelia; Surf Rock; Enstrümantal
Western Europe|Pop/Şarkı-Yazarı; Blues|Melankoli|Gece|İngiliz Folk-Blues; Manchester
Western Europe; North Africa; Fusion|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Gece|Tunisian Song; Acoustic Soul; Folk-Blues
Western Europe; North America|Blues; Caz|Romantik|Yolculuk|Harmonica; Vocal Jazz; Tenor Saxophone
Balkans|Vokal/Şarkı|Dinginlik|Gece|Yunan Entekno; Şiirsel Şarkı; Sanat Müziği
Western Europe|Caz|Neşe|Dans; Spor|Vintage Swing; Caz Vokal; Retro Bando
North America|Vokal/Şarkı; Caz|Melankoli|Gece|Kabare Vokali; Çok Dilli Repertuvar
Balkans|Vokal/Şarkı; Gelenek|Melankoli|Gece|Yunan Şarkısı; Plessas Bestesi
Anatolia|Elektronik; Caz|Karanlık|Odak; Yolculuk|Caz; Elektronik; Nublu Sahnesi
Western Europe|Gelenek|Dinginlik|Dans; Yolculuk|Çok Dilli Folk; Chanson; Ada Müzikleri Füzyonu
Anatolia|Elektronik|Karanlık|Odak|Anadolu Psychedelia; Elektro Saz; Surf-Funk
North America; Iranian World; West Africa; Fusion|Klasik/Modern Klasik|Dinginlik|Dans; Odak; Yolculuk|Fars-Ortaçağ Füzyonu; Kora; Erken Müzik
Balkans|Pop/Şarkı-Yazarı|Romantik|Gece|Yunan Indie Pop; Akustik Şarkı Yazımı
North America|Psikedelik/Rock|Dinginlik|Odak|Çöl Rock; Tex-Mex/Mariachi Nefeslileri; Tucson
Anatolia|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Anadolu Psychedelia; Surf Rock; Enstrümantal
Anatolia|Psikedelik/Rock|Karanlık|Odak; Yolculuk|Anadolu Psychedelia; Surf Rock; Enstrümantal
Caucasus|Reggae/Dub; Vokal/Şarkı|Dinginlik|Gece|Gürcü Rock/Reggae; 33A Topluluğu
Balkans|Gelenek|Melankoli|Odak; Gece; Yolculuk|Girit Halk Müziği; Girit Lyrası; Entehno
West Africa; North Africa; Fusion|Psikedelik/Rock|Melankoli|Yolculuk|Tuareg Rock; Assouf; Çöl Blues
Balkans|Vokal/Şarkı|Melankoli|Gece|Sevdalinka; Boşnak Halk Şarkısı
Western Europe|Tel/Enstrümantal|Epik|Odak; Gece|Enstrümantal; Doğu Avrupa Füzyon; Sinemasal Folk
Balkans|Nefesli/Bando|Neşe|Yolculuk|Bulgar Roman Düğün Müziği; Saksafon; Wedding Band
Western Europe|Elektronik|Neşe|Dans|Downtempo; Trip-Hop; Vintage Sample
North America|Pop/Şarkı-Yazarı|Epik|Dans; Yolculuk; Spor|Dene/Yerli Halk Anlatısı; Yellowknife
East & Central Asia; North America; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Guzheng; Banjo; Appalachian/Old-Time; Çin ve Uygur Halk Müziği; Kültürlerarası Düet
Iranian World|Caz|Melankoli|Yolculuk|Afgan Halk Müziği; Deneysel; Politik Caz
Northern Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Spor|İskandinav Şarkı-Yazarı; Orkestral Düzenleme
Iranian World|Tel/Enstrümantal|Romantik|Dans; Odak; Yolculuk|Nuevo Flamenco; İspanyol Gitar; Rumba
Anatolia|Klasik/Modern Klasik|Dinginlik|Odak; Yolculuk|Ney; Osmanlı Klasik Müziği; Sufi Gelenek
Iranian World|Vokal/Şarkı; Gelenek|Melankoli|Gece|İran Kadın Vokal İkilisi; Fars Şiiri
Southeast Asia|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Odak; Spor|60'lar-70'ler Saigon Soul Canlanması; Ho Chi Minh
Eastern Europe|Gelenek|Karanlık|Yolculuk|Ethno-Chaos; Ukrayna Halk Müziği; Avant-Folk
Southern Africa; South Asia; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak; Yolculuk|Acoustic Guitar; Indian Percussion; Double Bass
Mediterranean Europe|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Salento Vokal; Balkan-Akdeniz Repertuvar; Ağıt
Western Europe; Eastern Europe; Fusion|Gelenek|Neşe|Yolculuk|Doğu Avrupa/Roman Repertuvarı; Fransız Topluluğu
Anatolia|Tel/Enstrümantal|Melankoli|Dans; Yolculuk; Spor|Mikrotonal Gitar; Anadolu Makamı; Enstrümantal Düet
Western Europe; Balkans; Fusion|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans|Balkan Beats; Electro Swing; Transglobal Bass; Klezmer; Asian Breaks
Balkans; Western Europe|Klasik/Modern Klasik|Neşe|Odak|Yaylı Dörtlü; Çağdaş Repertuvar; Londra
Mediterranean Europe|Funk/Soul|Karanlık|Odak|İtalyan Film Müziği; Library Funk; Enstrümantal
Northern Europe|Psikedelik/Rock; Vokal/Şarkı|Karanlık|Yolculuk|İskandinav Rock; Karanlık Bariton
West Africa; Levant; Fusion|Blues; Tel/Enstrümantal|Epik|Odak; Yolculuk; Spor|Desert Blues; Acoustic Piano; Improvisation
North America|Caz; Vokal/Şarkı|Melankoli|Yolculuk|Caz Vokali; Blues/Bossa Etkisi; Philadelphia
Western Europe|Funk/Soul|Dinginlik|Dans; Odak; Yolculuk; Gece|Afro-Caz; Library Funk; Enstrümantal Groove
Caribbean|Reggae/Dub; Vokal/Şarkı|Neşe|Dans|Roots Reggae; Rocksteady; Acoustic Session
North America|Gelenek; Blues|Neşe|Gece|Siyah Telli Çalgı Geleneği; Banjo/Keman; Pulitzer
Western Europe|Psikedelik/Rock; Blues|Neşe|Dans; Yolculuk; Spor|Led Zeppelin Sonrası; Çöl Bluesu İşbirlikleri
North Africa; Western Europe; Fusion|Vokal/Şarkı; Caz|Neşe|Dans; Yolculuk; Spor|Arabic Poetry; Experimental Jazz; Oud/Brass
Western Europe|Gelenek|Neşe|Dans; Gece; Yolculuk|English Folk; Fiddle; Guitar; Electronic Folk
Western Europe; Levant|Tel/Enstrümantal; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Oud Ensemble; Arabian Contemporary; Instrumental
Levant|Caz|Neşe|Dans; Odak; Yolculuk|Arap Cazı; Piyano Üçlüsü; Makam-Swing
Eastern Europe; North America|Caz; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Gitar Cazı; Macar Ezgi Hattı
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Bariton Saksofon/Klarnet; ECM; Solo Çok Katmanlı
North America; South Asia; Fusion|Blues; Klasik/Modern Klasik|Melankoli|Odak; Gece; Yolculuk|Slayt Gitar + Mohan Veena; Grammy'li Buluşma
North Africa|Vokal/Şarkı|Melankoli|Dans|Sahrawi Repertuvarı; Sürgün/Direniş Şarkıları
Western Europe; South Asia; Fusion|Elektronik; Klasik/Modern Klasik|Melankoli|Dans|Asian Underground; Hint Klasiği + Elektronik; Londra
North America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Gece|Şair-Şarkı Yazarı; Alt Tonlu Vokal; Montreal
Northern Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece|Enstrümantal Gitar; Latin Füzyon; Downtempo
Latin America|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Tropikal Psikedeli; Deneysel Cumbia; Bogotá
North America|Caz|Karanlık|Odak; Spor|Jazz Fusion; Indian Classical Influences
Western Europe; Mediterranean Europe; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Dans|Kadın Vokal Topluluğu; Akdeniz Polifonisi; Paris
Eastern Europe|Reggae/Dub|Neşe|Yolculuk; Spor|Rus Halk Müziği; Folk-Rock; Ska Etkisi
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Gece|Gal Halk Müziği; Galçe Söz
Western Europe|Gelenek; Psikedelik/Rock|Melankoli|Yolculuk|İngiliz Folk-Rock; Maddy Prior; 1969
East & Central Asia|Gelenek; Psikedelik/Rock|Melankoli|Dans; Yolculuk; Spor|Moğol Folk-Rock; Gırtlak Vokali; Pekin
Northern Europe|Caz|Dinginlik|Odak; Gece|Karanlık Exotica; 1960’lar Latin Cazı; Vibrafon/Kontrbas/Perküsyon
Levant|Elektronik; Tel/Enstrümantal|Karanlık|Dans; Odak; Yolculuk; Spor|Ud/Saz + Synth; Makam; Osmanlı Halk Müziği; Psych/Trâns; Tel Aviv
West Africa|Vokal/Şarkı|Melankoli|Gece|Bamana Repertuvarı; Ngoni + Elektrikli Gitar; Bağımsız Söz
North Africa|Pop/Şarkı-Yazarı|Melankoli|Gece|Tunus Folk-Blues; Akustik Düet; Darija Şarkı
Latin America|Vokal/Şarkı|Neşe|Yolculuk|Slowcore Chanson; Meksika Bolero; Alternatif Folk
West Africa; Western Europe|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Bambara Vokali + Avrupa Prodüksiyonu; Paris
Western Europe; Iberia|Klasik/Modern Klasik; Vokal/Şarkı|Epik|Gece|Baroque; Early Music; Improvised Arrangement
East & Central Asia; Northern Europe; Western Europe; Fusion|Caz; Vokal/Şarkı|Melankoli|Gece|Çağdaş Avrupa Cazı; Caz Vokali; Gitar; Bas/Çello; Perküsyon; Doğaçlama
Levant|Reggae/Dub|Melankoli|Dans; Odak; Yolculuk|Reggae; Çöl Blues; Arap Rock
Western Europe|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Klezmer Klarneti; 'Yeni Yahudi Müziği'; Paris
Latin America|Klasik/Modern Klasik|Neşe|Gece|Exotica; And Etkili Vokal; Orkestral Pop
Levant|Caz|Dinginlik|Dans; Odak; Gece; Yolculuk|Çeyrek Tonlu Trompet; Arap Makamı + Caz
North America; West Africa; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Free Jazz; Mali Korası; Mande Geleneği; Doğaçlama
Northern Europe|Caz; Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk; Spor|Nordic Jazz; Progressive Rock; Folk; Ambient; Improvisation
Western Europe; Latin America; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece|Piano/Percussion Duo; Argentine Folk; Jazz
Western Europe|Psikedelik/Rock|Neşe|Dans|Mestizo Rock; Latin-Reggae Füzyon; Politik Şarkı
Eastern Europe|Gelenek|Melankoli|Gece|Macar Táncház Hareketi; Transilvanya Alan Kayıtları
Iberia; Western Europe; Fusion|Pop/Şarkı-Yazarı; Elektronik|Neşe|Yolculuk|Flamenco Fusion; Electronica; Global Beats
Latin America|Vokal/Şarkı; Elektronik|Neşe|Yolculuk|Çağdaş Bossa Nova; Elektronik Prodüksiyon
Western Europe; West Africa|Blues; Tel/Enstrümantal|Neşe|Dans; Spor|Çöl Bluesu Gitarı; Prodüktörlük; Robert Plant/Tinariwen
Western Europe; South Asia; Fusion|Elektronik; Klasik/Modern Klasik|Neşe|Yolculuk|Asian Underground; Hint Klasiği + Elektronik; Londra
Western Europe; Fusion|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Spor|Balkan Beats; Electro Swing; Transglobal Bass; Klezmer Etkileri
Anatolia; Northern Europe|Caz; Gelenek|Romantik|Yolculuk; Spor|Turkish Clarinet; Contemporary Jazz; Live Session
Western Europe|Elektronik|Melankoli|Gece|Klezmer Füzyon; Balkan; Elektronik Folk
West Africa|Blues|Romantik|Dans; Gece; Yolculuk|Mali Blues; Afro-Pop; Gitar Düeti
Western Europe; North America; Fusion|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Gece|Art Pop; Film Music; Experimental Song
East Africa|Funk/Soul; Vokal/Şarkı|Neşe|Yolculuk|Ethio-Soul; Éthiopiques; Addis
North Africa; Caribbean; Fusion|Gelenek; Caz|Neşe|Gece; Yolculuk|'Pianoriental'; Oran; Endelüs + Kuba
West Africa|Tel/Enstrümantal; Gelenek|Neşe|Yolculuk|Ngoni; Mande Geleneği; Ngoni Ba
Caribbean|Gelenek; Vokal/Şarkı|Neşe|Dans|Garifuna; Paranda; Punta
Western Europe|Gelenek|Neşe|Dans|İngiliz Halk Müziği; Çok Kültürlü Kolektif; Elektronik Folk
Latin America|Tel/Enstrümantal|Neşe|Dans; Odak; Spor|Akustik Gitar Düeti; Flamenko-Metal Füzyon
East & Central Asia|Vokal/Şarkı; Elektronik|Epik|Spor|Çin Vokali; Tibet/Moğol Esini; Film Müziği
North America|Gelenek|Melankoli|Dans|İlerici Bluegrass; Chris Thile; Mandolin/Banjo
Western Europe|Pop/Şarkı-Yazarı; Reggae/Dub|Neşe|Dans|Çağdaş Chanson; Reggae/Ska; Toplumsal Söz
North America|Blues; Pop/Şarkı-Yazarı|Neşe|Dans|Tulsa Sound; Sade Blues-Rock; Oklahoma
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Piyano Üçlüsü; Çağdaş Caz; Toulouse
West Africa|Blues|Neşe|Yolculuk|Kamele Ngoni; Mande Müziği; Çöl Blues
West Africa|Vokal/Şarkı|Romantik|Dans|Morna; Coladeira; Çağdaş Cape Verde Pop
East Africa|Vokal/Şarkı|Melankoli|Odak; Gece; Yolculuk|Acholi Geleneği; Sürgün Şarkıları; Real World
North America; Latin America; West Africa; Fusion|Funk/Soul; Psikedelik/Rock|Neşe|Odak|Enstrümantal Gitar; Chicha/Tuareg/Son Montuno; New Orleans
North America|Funk/Soul|Epik|Dans; Odak; Spor|Afrobeat; Ethio-Jazz; Politik Funk
West Africa|Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece|Cavaquinho/Gitar; Yeşil Burun Geleneği
West Africa; Latin America; Fusion|Elektronik; Gelenek|Neşe|Dans|Afro-Folk; Worldbeat; Experimental Pop
Iberia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans; Gece|Fado Etkili Çağdaş Portekiz Pop'u; Lizbon
West Africa; Caribbean; Fusion|Gelenek|Romantik|Dans; Yolculuk|Afro-Küba Füzyonu; Mande Müziği; Son Cubano
North America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|Americana; Mandolin/Banjo; Teksas
Eastern Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Dans; Odak; Yolculuk|Klezmer; Yahudi Kamara Müziği; Doğaçlama
Western Europe; West Africa; Fusion|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Dans|Malian Kora; Afro-Pop; French Pop; Manding Vocals
Iberia|Vokal/Şarkı|Melankoli|Dans|Çağdaş Fado; Lizbon; Fado Ailesi
East & Central Asia|Gelenek|Melankoli|Dans|Moğol Gırtlak Şarkıcılığı; Morin Khuur; İç Moğolistan
North America|Gelenek; Tel/Enstrümantal|Dinginlik|Gece|Banjo Düeti; Appalachian/Old-Time
Latin America|Vokal/Şarkı; Gelenek|Neşe|Dans; Spor|Carimbó; Amazon/Pará Repertuvarı
West Africa|Blues; Psikedelik/Rock|Neşe|Dans; Spor|Çöl Blues; Timbuktu'dan Sürgün; Bamako
Northern Europe|Gelenek; Elektronik|Karanlık|Odak; Spor|Genişletilmiş Akordeon; Deneysel; Helsinki
Latin America|Gelenek|Neşe|Dans|Cumbia; Bullerengue; Afro-Kolombiya Perküsyonu
West Africa|Psikedelik/Rock|Epik|Dans; Spor|Tuareg Rock; Çöl Blues; Assouf
Caribbean|Caz|Dinginlik|Odak; Yolculuk; Spor|Kuba Caz Piyanosu; Buena Vista Bağlantısı; Afro-Küba
Caribbean|Vokal/Şarkı; Funk/Soul|Neşe|Dans|Calypso; Soca; Kadin Öncülük
Northern Europe|Vokal/Şarkı; Gelenek|Epik|Dans; Yolculuk|Sami Joik; Yerli Halk Direnişi
West Africa; Southern Africa; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece|Fula Şarkısı; Akustik Gitar; Mbalax Kökleri
East & Central Asia|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Geomungo + Daegeum; Kore Doğaçlaması; Seul
West Africa; North America; Fusion|Gelenek; Klasik/Modern Klasik|Neşe|Dans; Yolculuk|Mande Griot; Balafon/Ngoni; String Quartet
West Africa|Psikedelik/Rock|Neşe|Yolculuk|Tuareg Rock; Çöl Blues; Assouf
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Dans; Yolculuk|Tropikália'nın Kurucusu; Bahia; MPB
Western Europe; East Africa; Fusion|Caz|Neşe|Odak; Yolculuk|Ethio-Jazz; Enstrümantal Funk
Levant|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Yemen-Yahudi Kadın Repertuvarı; Arapça Vokal; Elektronik Prodüksiyon
Western Europe; North America|Caz; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece; Yolculuk|Gypsy Jazz; Manouche Guitar; Mandolin; Contemporary Jazz
Caucasus|Caz; Klasik/Modern Klasik|Neşe|Odak; Yolculuk|Jazz Piano; Armenian Modalism; Contemporary Composition
Levant; Anatolia; Iranian World; Fusion|Elektronik; Gelenek|Neşe|Odak|Ethnic Electronica; Kamancha; Saz; Dub; Global Dance
Western Europe|Elektronik|Neşe|Dans|Balkan Brass; Elektronik; Sokak Bandosu
Western Europe; West Africa; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Kelt Arpı + Kora İkilisi
Levant|Elektronik|Epik|Dans; Yolculuk; Spor|Shamstep; Dabke; Elektro-Arap
Mediterranean Europe|Klasik/Modern Klasik; Gelenek|Epik|Dans|Pizzica; Minimalist Piyano; Salento Folk
Levant; Western Europe; Fusion|Tel/Enstrümantal; Caz|Melankoli|Odak; Gece|Kanun + Çello + Perküsyon Üçlüsü
Western Europe|Vokal/Şarkı|Melankoli|Gece; Yolculuk|İngiliz Halk Müziği; Keman-Vokal; Aile Geleneği
Southern Africa|Gelenek; Funk/Soul|Neşe|Yolculuk|Soweto; Ritmüel Vokal/Perküsyon; Uzun Formlu İcra
South Asia|Klasik/Modern Klasik|Romantik|Dans; Gece|Hindustan Klasik; Slide Gitar; Aile Topluluğu
West Africa|Tel/Enstrümantal|Neşe|Dans; Yolculuk|Mande Gitarı; Super Rail Band; Bamako
North America|Funk/Soul|Neşe|Dans|Müzikal Tiyatro; Funk-Soul; Broadway Besteci
Caribbean|Vokal/Şarkı|Neşe|Dans|Haiti Folk; Alternatif Gitar; Kreol Şarkı
East Africa; North America|Vokal/Şarkı; Gelenek|Dinginlik|Dans; Yolculuk|Sudan Retro-Pop'u; Nübye Repertuvarı; Brooklyn
Mediterranean Europe|Nefesli/Bando|Epik|Dans; Odak|Sicilya Bandası; Tören/Alay Repertuvarı
Balkans|Vokal/Şarkı|Romantik|Dans; Gece|Rebetiko; Sefarad/Judeo-İspanyol; Balkan Repertuvarı
Western Europe; Central Africa; Fusion|Gelenek|Neşe|Dans; Odak; Yolculuk|Baka Polifonisi + Kelt Hattı; Real World
Latin America|Elektronik; Vokal/Şarkı|Dinginlik|Yolculuk|Dijital Cumbia + Vokal
Western Europe|Funk/Soul; Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Neo-Soul; Gitar-Vokal; Yunan-Jamaika Kökeni; Londra
Western Europe|Gelenek|Neşe|Dans; Spor|Gal Halk Müziği; Step Dans; Arp/Keman
Latin America|Elektronik|Karanlık|Yolculuk|Cumbia; Elektronik Pop; Yerli Kimlik
Western Europe; Anatolia; Fusion|Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Balkan Elektro; Bucovina Club; Frankfurt
Central Africa|Blues|Melankoli|Dans; Gece|Bassa Folk; Minimal Akustik; Afro-Blues
Caribbean|Tel/Enstrümantal|Melankoli|Gece|Haiti-Louisiana Füzyon; Çello-Banjo; Kreol Şarkı
West Africa|Funk/Soul|Neşe|Dans; Yolculuk; Spor|Gana Afro-Funk; Brian Eno Prodüksiyonu; 1981
Balkans|Vokal/Şarkı; Gelenek|Melankoli|Gece|Sevdalinka; Saz/Gitar; Saraybosna
Iberia; Western Europe; Fusion|Vokal/Şarkı; Caz|Melankoli|Gece|Bossa-Inspired Jazz; Portuguese Song; Guitar Dialogue
West Africa; Western Europe; North Africa; Fusion|Blues; Gelenek|Dinginlik|Dans; Odak; Gece; Yolculuk|Malian Guitar; Desert Blues; Moroccan Violin
Western Europe|Elektronik; Caz|Dinginlik|Dans; Odak; Gece|Nu-Jazz/Downtempo; Sonar Kollektiv; Berlin
North America|Blues; Vokal/Şarkı|Romantik|Gece; Yolculuk|Piyano-Blues; New Orleans Kökeni; Paris
Latin America|Pop/Şarkı-Yazarı|Neşe|Dans|Latin Folk-Pop; İki Dilli Şarkı; Gitar
North America|Funk/Soul; Caz|Melankoli|Yolculuk|Folk Soul; Chicago Soul; Jazz Folk
Central Africa|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Kongo Rumbası; Soukous; TPOK Jazz
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Fransız Alternatif Chanson; Minimal Düzenleme
Caribbean|Vokal/Şarkı; Gelenek|Melankoli|Gece|Karayip Gitar-Vokal; Calypso/Kreol Repertuvar
Eastern Europe|Gelenek; Caz|Neşe|Dans; Yolculuk; Spor|Doğu Avrupa-Kafkas Halk Müziği; Panduri; Prog-Caz Füzyon
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece|İrlanda Geleneksel Şarkısı; Ulster/Tyrone; İrlanda ve İskoç Gal Repertuvarı
Western Europe|Gelenek|Melankoli|Dans; Gece|Bluegrass; Aile Topluluğu; Kuzey İrlanda
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans|Celtic Folk; Roots Americana; Banjo; Guitar; Vocal Harmony
Iranian World; Western Europe|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Fars Vokali + Elektronik; Ritual Estetigi
East Africa|Vokal/Şarkı; Caz|Epik|Gece|Ethio-Jazz + Amerikan Cazı; San Francisco
East Africa|Gelenek; Funk/Soul|Neşe|Dans; Yolculuk|Sudanese Jazz; Sudanese Pop; Funk; Rock
Levant|Klasik/Modern Klasik|Epik|Odak; Yolculuk|Oud Üçlüsü; Arap Klasik Müziği
Western Europe|Blues; Elektronik|Romantik|Dans; Yolculuk|Blues-Gospel + Dans Ritimleri; Amsterdam
Mediterranean Europe; Iberia; Latin America; Fusion|Klasik/Modern Klasik; Tel/Enstrümantal|Dinginlik|Odak; Gece|Classical Guitar; Latin American Repertoire; Spanish Guitar
North America|Pop/Şarkı-Yazarı|Melankoli|Odak; Gece; Yolculuk|Amerikan Indie Folk; Anlatıcı Söz
Western Europe|Psikedelik/Rock; Gelenek|Melankoli|Yolculuk|İngiliz Folk-Rock'ın Kurucuları; 1967
Western Europe; Caribbean|Blues|Neşe|Dans; Spor|Guadeloupe Bluesu; Kreol Söz; Dobro + Tuba
Western Europe|Elektronik; Klasik/Modern Klasik|Dinginlik|Dans; Odak; Gece; Yolculuk|Viyola da Gamba + Elektronik; Ambient
Anatolia|Gelenek; Psikedelik/Rock|Neşe|Dans|Anatolian Pop; Bektaşi Deyişi; Rebetiko; Psychedelic Rock
Levant|Caz; Gelenek|Neşe|Odak|Piano Trio; West African Inspiration; Jazz Improvisation
Western Europe; Fusion|Elektronik; Reggae/Dub|Neşe|Dans|Mediterranean Electronica; Dub; Jungle; Gnawa; Sephardic Folk
Iberia|Gelenek|Neşe|Yolculuk; Spor|Flamenko; Utrera Ekolü; Kuba Tres Gitarı
North America|Pop/Şarkı-Yazarı|Romantik|Gece; Yolculuk|Alt-Country; Piano Ballad; Guest Vocal
Western Europe|Caz; Tel/Enstrümantal|Epik|Dans; Odak; Yolculuk|Beş Telli Kontrbas; Flamenko/Akdeniz; Enstrümantal
Western Europe; Levant|Pop/Şarkı-Yazarı; Vokal/Şarkı|Dinginlik|Odak; Gece; Yolculuk|Indie Pop; Folk-Pop; Piano; Multilingual Song
Eastern Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Klezmer; Yahudi Kamara Müziği; Doğaçlama
Western Europe; North America|Pop/Şarkı-Yazarı; Funk/Soul|Dinginlik|Yolculuk; Spor|Psychedelic Soul; Cover Album; Vocal Duet
Western Europe; Central Africa; Fusion|Funk/Soul|Dinginlik|Gece|Soul; Kamara Pop; Afro-Fransız İşbirliği
Latin America|Pop/Şarkı-Yazarı; Reggae/Dub|Neşe|Gece|Brezilya Popı; Reggae/Baile Funk; Paris
Western Europe|Gelenek|Melankoli|Dans; Gece|Chanson; Indie Folk; Piyano Şarkısı
North America|Pop/Şarkı-Yazarı; Funk/Soul|Neşe|Dans|Soul-Folk; Be Good Tanyas Çevresi; Vancouver
North America|Funk/Soul; Caz|Melankoli|Dans|Soul-Caz Vokali; Haiti Kökeni; Montreal
North America|Gelenek; Vokal/Şarkı|Neşe|Dans|Klezmer; Yiddish Song; Jewish Roots
Balkans|Nefesli/Bando; Elektronik|Neşe|Dans; Odak; Spor|Balkan Brass; Electronic Remix; Dance
Iranian World|Vokal/Şarkı; Gelenek|Melankoli|Gece|Fars Vokali; Fars Şiiri; Diaspora Kayıtları
Mediterranean Europe; Northern Europe; Fusion|Gelenek; Elektronik|Neşe|Dans; Spor|Éntekhno; Cypriot Folk; Electronic
North Africa; Oceania; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak; Gece; Yolculuk|Ud Virtüozluğu; Arap Makamı + Caz; Sydney
Latin America|Vokal/Şarkı; Elektronik|Neşe|Yolculuk|Tango; Rioplatense Song; Electronic Production; String Ensemble
Eastern Europe|Gelenek|Melankoli|Gece; Yolculuk|Ethno-Chaos; Ukrayna Halk Müziği; Avant-Folk
North America|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak|Bluegrass; Newgrass; Fiddle
South Asia; Eastern Europe; Fusion|Klasik/Modern Klasik; Tel/Enstrümantal|Neşe|Odak|Carnatic Violin; Romani Violin; Tabla; Jazz Fusion
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Fransız Popu; Şanson
North America|Caz|Neşe|Dans; Yolculuk|Caz Vokal; Soul; Broadway Geleneği
Iranian World; North America|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Gece|Farsça İndie-Pop; Animasyon; Vancouver
Western Europe|Caz|Melankoli|Dans; Odak; Yolculuk|Gypsy Jazz; Manouche Swing; Akustik Gitar Topluluğu
East Africa|Gelenek; Vokal/Şarkı|Melankoli|Dans|Maloya; Réunion Kreolü; Çağdaş Akustik/Dünya Müziği
Levant; Caucasus; Fusion|Tel/Enstrümantal|Epik|Odak|Kamançe; Dağıstan/Kafkas Kökeni; Ortadoğu Enstrümantali
Western Europe|Elektronik|Melankoli|Dans; Odak|Balkan Elektro; Bucovina Club; Frankfurt
Anatolia; North America|Elektronik; Gelenek|Epik|Dans; Odak; Yolculuk|Ney + Elektronik; Sufi Esini; Arkın Allen
North America; North Africa; Fusion|Elektronik; Reggae/Dub|Melankoli|Dans; Odak|Electrocumbia; Dub; Global Bass
Western Europe|Hip-Hop/Rap|Neşe|Dans; Odak; Spor|Trip-Hop; Sample Bazlı Enstrümantal Hip-Hop
North America|Blues; Pop/Şarkı-Yazarı|Neşe|Dans|Experimental Folk; Blues; Old-Time; Gypsy Jazz; Vocal Layering
Iberia|Caz|Neşe|Dans; Odak|Gypsy Jazz; Manouche Swing; Akustik Gitar; Mallorca
North America; Levant; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece|Frame Drum; Oud; Ney; Kamancheh; Ambient Healing Music
North Africa|Psikedelik/Rock|Neşe|Dans|Gnawa Rock; Psikedelik; Guembri Füzyon
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Çağdaş Şanson; Nancy/Paris
Western Europe|Tel/Enstrümantal|Neşe|Dans; Odak|Enstrümantal; Doğu Avrupa Füzyon; Sinemasal Folk
Eastern Europe; North America|Elektronik; Vokal/Şarkı|Dinginlik|Odak|Downtempo/Trip-Hop + Şarkı-Yazarı Vokali
Mediterranean Europe; Fusion|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Greek Art Song; Archival Vocal; Electronic Rework
North America|Blues; Vokal/Şarkı|Melankoli|Gece|50'ler R&B; Kadın Blues Vokali
Levant; Iberia; Fusion|Vokal/Şarkı; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Flamenco Vocal; Guitar; Middle Eastern Ensemble
Eastern Europe; Northern Europe; Fusion|Caz; Gelenek|Neşe|Dans; Odak|Bulgar Halk Müziği; Finlandiya Cazı; Aksak Ritimler; Sinematik Düzenleme
North America|Caz; Tel/Enstrümantal|Karanlık|Odak; Yolculuk|Caz Gitarı; Americana Etkisi; ECM/Nonesuch
Northern Europe|Tel/Enstrümantal; Gelenek|Neşe|Dans; Spor|Akustik İkili; Halk Müziği; Funk/Rock/Caz Etkileri
Western Europe|Caz; Psikedelik/Rock|Karanlık|Yolculuk|Jazz Rock; Spoken Word; Poetry
Anatolia|Vokal/Şarkı|Epik|Dans; Yolculuk; Spor|Kürt Halk Müziği; Anadolu Vokal Geleneği
Anatolia|Vokal/Şarkı|Romantik|Yolculuk|Kürt Halk Müziği; Anadolu Vokal Geleneği
Levant|Reggae/Dub|Karanlık|Dans; Odak; Yolculuk; Spor|Reggae; Çöl Blues; Arap Rock
North America; Fusion|Caz; Elektronik|Dinginlik|Odak; Gece|Jazz Fusion; Piano Improvisation
North America|Psikedelik/Rock; Pop/Şarkı-Yazarı|Dinginlik|Odak; Yolculuk|Indie Rock; Noise Pop; Dream Pop; Experimental Rock
Western Europe|Elektronik|Neşe|Odak; Spor|Enstrümantal Psikedelik; Uzakdoğu Etkili Funk; Kraut-Disco
West Africa|Blues; Psikedelik/Rock|Neşe|Dans; Spor|Çöl Blues; Timbuktu'dan Sürgün; Bamako
North America|Reggae/Dub|Neşe|Spor|Cinematic Soul; Enstrümantal Funk; Dub-Soul
South Asia; Western Europe; Fusion|Elektronik|Neşe|Dans; Yolculuk; Spor|Asian Underground; Tabla + Drum'n'Bass; Londra
Latin America|Elektronik|Melankoli|Odak; Gece|Deneysel Elektronik; Other People; New York-Santiago
Levant|Funk/Soul|Neşe|Dans|Enstrümantal Funk; Sample Prodüksiyon; Konuk Vokal
East Africa|Funk/Soul; Gelenek|Neşe|Yolculuk; Spor|Cibuti Devlet Radyo-TV Orkestrası; Somali/Etiyopya Hattı
Caribbean|Caz|Romantik|Odak; Gece|Keman-Vokal; Afro-Küba Cazı; Santería Etkisi
Eastern Europe|Nefesli/Bando; Elektronik|Neşe|Dans; Spor|Balkan-Roman Nefesli; Budapeşte; Kulüp Ritimleri
Eastern Europe|Gelenek|Melankoli|Gece|Polish Folk; Polyphonic Vocals; Object Percussion
North America|Vokal/Şarkı|Dinginlik|Gece|Sefarad ve Mizrahi Repertuvar; Kadın Vokal Topluluğu
Levant; Arabian Peninsula|Vokal/Şarkı; Elektronik|Neşe|Yolculuk|Yemen-Yahudi Kadın Repertuvarı + Elektronik
North America|Vokal/Şarkı; Gelenek|Melankoli|Gece|Doom-Folk; İrlanda Halk Ezgileri; Keman
Western Europe; Anatolia; Fusion|Psikedelik/Rock; Elektronik|Karanlık|Odak; Spor|Belçika Groove + Elektro Saz (BaBa ZuLa Hattı)
North America; West Africa; Fusion|Caz; Gelenek|Neşe|Dans; Yolculuk; Spor|Wassoulou; West African Rhythm; Jazz Ensemble
Western Europe; South Asia; Caribbean; Fusion|Caz; Gelenek|Neşe|Dans; Yolculuk|Bengal Halk Ezgileri + Kuba Ritimleri; Londra
Western Europe|Elektronik; Pop/Şarkı-Yazarı|Epik|Dans; Odak; Yolculuk; Spor|Dub; Worldbeat; Electronic Pop; Cross-Cultural Fusion
West Africa; North America; Fusion|Gelenek; Vokal/Şarkı|Neşe|Odak|West African Folk; Appalachian Banjo; Accordion
Western Europe|Caz; Funk/Soul|Neşe|Odak; Yolculuk|UK Jazz; Afrobeat; Psychedelic Groove; Spiritual Jazz
Northern Europe|Caz|Dinginlik|Odak|Enstrümantal; Kütüphane Müziği; Anadolu-Japon Etkili Caz-Funk
Northern Europe|Caz|Dinginlik|Odak|Enstrümantal; Kütüphane Müziği; Anadolu-Japon Etkili Caz-Funk
Latin America; East & Central Asia; Fusion|Gelenek|Neşe|Dans; Yolculuk|Cumbia; Japon Min'yo; Tropikal Füzyon
North America|Reggae/Dub|Neşe|Dans; Odak; Spor|Psikedelik Funk; Tay-Dub Etkili Enstrümantal
North America|Reggae/Dub|Neşe|Dans; Odak; Spor|Psikedelik Funk; Tay-Dub Etkili Enstrümantal
Balkans|Elektronik; Gelenek|Neşe|Dans; Odak; Spor|Yunan Elektroniği + Kanun
Western Europe|Gelenek|Karanlık|Dans; Odak|İngiliz Geleneksel Şarkı; Ham Prodüksiyon; Londra
Western Europe; Mediterranean Europe|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Downtempo; Trip-Hop; Vintage Sample; Electro Swing; Vintage Pop
Anatolia; North America; Iberia; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Ottoman-Turkish Music; Sephardic History; Contemporary Jazz
Levant|Elektronik|Karanlık|Dans; Odak; Spor|Arap Elektronik; Tarab Füzyon; Vokal + Prodüksiyon
Western Europe|Caz|Epik|Odak; Yolculuk|Çello; Caz-Klasik Füzyon; Prodüksiyon
Iranian World; Oceania; Fusion|Caz; Gelenek|Dinginlik|Odak; Yolculuk|Persian Tar; Contemporary Jazz; Ensemble Improvisation
Iranian World; Oceania|Gelenek; Caz|Melankoli|Gece; Yolculuk|Kamançe/Qeychak + Vokal; Melbourne
Western Europe|Reggae/Dub; Elektronik|Neşe|Dans; Odak|Roots Reggae; Dub; Dancehall; Hip-Hop
South Asia; Eastern Europe; Fusion|Klasik/Modern Klasik; Tel/Enstrümantal|Neşe|Dans; Odak; Yolculuk; Spor|Carnatic Violin; Romani Violin; Tabla; Jazz Fusion
Western Europe; Balkans; Fusion|Gelenek; Nefesli/Bando|Neşe|Odak|Klezmer + Balkan/Romen Repertuvarı
Mediterranean Europe|Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|İtalyan Indie; Folk-Blues; Vokal-Gitar Düeti
West Africa; Levant; Fusion|Blues; Tel/Enstrümantal|Dinginlik|Odak|Desert Blues; Acoustic Piano; Improvisation; Piyano-Gitar Doğaçlaması; İsrail-Mali Füzyonu; İsrail-Mali Füzyon; Piyano-Gitar Düeti; Guitar; West African Groove
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Saha Kayıtları; Deneysel Elektronik; Java Etkisi
West Africa|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk|Kadın Askeri Orkestra; Mande Repertuvarı; Konakri
Northern Europe|Psikedelik/Rock|Karanlık|Gece; Yolculuk|Tek Kişilik Blues; Garage; Dark Folk
Western Europe|Elektronik|Melankoli|Gece; Yolculuk|Downtempo; Electro Swing; Chillout Prodüksiyon
Western Europe|Elektronik; Klasik/Modern Klasik|Melankoli|Odak; Gece|Air'in Yarısı; Film Müziği; Paris
West Africa|Caz; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Afro-Swing; Saksofon + Vokal; Konakri
Eastern Europe; North America|Caz; Tel/Enstrümantal|Melankoli|Yolculuk|Gitar Cazı; Macar Ezgi Hattı
North Africa|Elektronik|Karanlık|Gece|Arap Alternatif; Protest Şarkı; Elektronik
West Africa|Blues; Psikedelik/Rock|Neşe|Dans; Spor|Çöl Blues; Timbuktu'dan Sürgün; Bamako
East & Central Asia|Gelenek; Psikedelik/Rock|Neşe|Yolculuk; Spor|Şaman Folk-Rock; Kore Geleneksel Çalgıları; Seul
Western Europe|Funk/Soul; Pop/Şarkı-Yazarı|Melankoli|Dans|Soul-Folk; Nijerya-Roman Kökeni
North Africa|Blues|Romantik|Odak; Yolculuk; Spor|Gnawa Füzyon; Blues; Soul
Western Europe|Gelenek|Melankoli|Yolculuk|Irish Traditional; Improvised Folk; Live Reactive Composition
Western Europe|Gelenek; Elektronik|Neşe|Yolculuk; Spor|Celtic Fusion; Acid Croft; Folk-Rock; Scotland
West Africa; North Africa; East Africa; Fusion|Klasik/Modern Klasik|Dinginlik|Dans; Odak; Gece; Yolculuk|Kora-Oud-Valiha Üçlüsü; Akustik Kamara Afrika
North America; Levant; Fusion|Tel/Enstrümantal|Neşe|Odak; Gece|Çello-Oud Düeti; Arap Makamı; Kamara Füzyon
Western Europe|Caz|Neşe|Odak|Çello; Caz-Klasik Füzyon; Prodüksiyon
West Africa|Pop/Şarkı-Yazarı|Neşe|Yolculuk|Afro-Küba; Senegal Pop; Mbalax Öncesi
North Africa|Vokal/Şarkı|Neşe|Dans|Sahrawi Müziği; Haul; Politik Şarkı
Oceania|Gelenek|Karanlık|Gece|Balkan-Akdeniz Füzyon; Çok Dilli Vokal; Kamara Folk
Balkans|Psikedelik/Rock; Funk/Soul|Neşe|Odak|Anadolu Rock Etkisi; Tay-Kamboçya Funk; Enstrümantal Psikedelia
South Asia|Klasik/Modern Klasik; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Hindustani Slayt Gitarı; Kalküta
West Africa|Blues|Melankoli|Odak|Tuareg Assouf; Çöl Blues; Elektrikli Gitar
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans|Fransız Şansonı; Politik Söz; Dünya Ritimleri
South Asia|Tel/Enstrümantal; Klasik/Modern Klasik|Dinginlik|Odak; Yolculuk|Sitar + Sarod + Perküsyon; Hindustani Çağdaşı
Caucasus|Gelenek|Neşe|Dans; Gece|Kafkas-Anadolu Ortak Repertuvarı; Duduk/Tel
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Dans; Yolculuk|Kelt Arpı; İskoç Gaelcesi Repertuvar; Glasgow
Northern Europe|Klasik/Modern Klasik; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Danish Art Song; Guitar; Recorder; Chamber Music
East & Central Asia|Gelenek|Melankoli|Dans|Moğol Gırtlak Vokali; Morin Khuur; Ulan Batur
Iberia; Mediterranean Europe; South Asia; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak|Indian Raga; Sarod; Tabla; Modal Music
Latin America; Iranian World; Fusion|Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Electrocumbia; Downtempo; Organic House; Dub
North Africa|Caz; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Ud; Sufi Vokal; Çağdaş Caz
North Africa|Elektronik|Epik|Dans; Yolculuk; Spor|Raï Rock; Elektronik Chaâbi; Punk
Western Europe|Caz|Epik|Dans; Yolculuk; Spor|Britanya Cazı; Karayip Ritmi; Politik Enstrümantal
Levant|Caz|Romantik|Odak; Gece; Yolculuk|Piyano; Arap Cazı; Şarkı Yazımı
West Africa|Blues|Melankoli|Odak; Gece|Bamako Blues; Elektrikli Gitar; Bambara Şarkısı
Northern Europe|Caz|Karanlık|Dans; Odak; Yolculuk|Nordik Caz Üçlüsü; Wesseltoft/Berglund/Öström
North America|Klasik/Modern Klasik|Melankoli|Odak|Deneysel Klasik; Perküsyon; Kanon-Round Formu
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Enstrümantal Elektronik; Asya-Ortadoğu Sample; Psikedelik
Western Europe|Caz|Dinginlik|Odak; Yolculuk|Çağdaş Fransız Cazı; Piyano Merkezli Caz; Piyano/Çello Dokusu
Western Europe|Psikedelik/Rock|Epik|Dans; Odak|Afro-Krautrock; Enstrümantal Funk; Etiyopya Etkisi
West Africa; Western Europe; Fusion|Gelenek; Elektronik|Neşe|Dans|Malian Griot Music; Kora; Field Recording; Dub Electronics
Balkans|Vokal/Şarkı|Melankoli|Dans|Rebetiko; Sefarad/Judeo-İspanyol; Balkan Repertuvarı
West Africa|Funk/Soul; Blues|Karanlık|Yolculuk; Spor|Blufunk; Perküsif Gitar; Lagos-Londra
North America; Iranian World; West Africa; Fusion|Klasik/Modern Klasik|Dinginlik|Odak; Gece; Yolculuk|Fars-Ortaçağ Füzyonu; Kora; Erken Müzik
East Africa|Funk/Soul; Vokal/Şarkı|Neşe|Yolculuk|Ethio-Soul; Swinging Addis; Éthiopiques
Western Europe|Reggae/Dub; Caz|Dinginlik|Dans; Odak|İngiliz Dub Prodüksiyonu + Retro Caz-Funk
Levant|Caz; Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Yolculuk|Lebanese Jazz; Arabic Song; Theatre Music; Piano; Political Satire
Southern Africa; Western Europe|Vokal/Şarkı; Gelenek|Melankoli|Dans|Angolan Song; Semba; Vocal Duet
East & Central Asia|Gelenek|Epik|Odak; Gece; Yolculuk|Tuva Gırtlak Şarkıcılığı; Igil/Doshpuluur
Latin America; Mediterranean Europe; West Africa; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans|Afro-Atlantik Repertuvar; Brezilya-Angola-İtalya Kadrosu
Latin America|Hip-Hop/Rap; Vokal/Şarkı|Neşe|Dans; Spor|Brazilian Rap; MPB; Nylon-String Guitar
Caucasus|Gelenek|Neşe|Dans|Ermeni Halk Müziği; Balkan Füzyon; Pop
Iberia|Elektronik|Melankoli|Dans; Gece|Asturya Folk; Queer Kabare; Deneysel Prodüksiyon
Oceania|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Folk-Rock; Sahne Kimliği Oyunları; Lyttelton
Latin America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Gece; Yolculuk|MPB; Pernambuco Hattı; Gitar-Vokal
Northern Europe; Latin America; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Yolculuk|Kutsal Dünya Müziği; Medicine Music; Yerli Gelenekler; Latin Amerika Halk Müziği; İskandinav Halk Müziği
Northern Europe|Tel/Enstrümantal|Neşe|Odak|Armónika Dörtlüsü; Fin Halk Müziği; Tango-Klezmer
Eastern Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Polonya Tarih Şarkıları; Retro Pop; Ukulele-Piyano
Western Europe; North Africa; Fusion|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Odak; Spor|Fransız Dünya Müziği; Tuareg/Mağrip İşbirlikleri; Angers
Southern Africa; Western Europe; West Africa; North America; Fusion|Funk/Soul; Elektronik|Karanlık|Dans; Yolculuk; Spor|Afrobeat; Electronic Production; Global Collective
Western Europe|Funk/Soul|Neşe|Dans; Odak; Spor|Deep Funk; Soul; Latin Funk
Iberia|Vokal/Şarkı|Melankoli|Dans; Gece|Katalan Vokali; Flamenko-Caz; Fado Etkisi
North America|Caz|Melankoli|Gece|Caz Vokal; Soul; Blues
Latin America|Vokal/Şarkı; Gelenek|Dinginlik|Dans; Yolculuk|Medicine Music; Deneysel Şarkıcı-Şarkı Yazarı; Törensel/Dünya Müziği
Western Europe|Reggae/Dub|Melankoli|Dans; Gece|Dub; Roots Reggae
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Samimi Folk; Şarkıcı-Şarkı Yazarı; Şanson; Atmosferik Bedroom Folk
Western Europe|Funk/Soul; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Psychedelic Soul; Analog Funk; Vocal Band
Western Europe|Elektronik; Klasik/Modern Klasik|Epik|Dans; Odak|Ethereal/Neoklasik; Erken Müzik; Ortadoğu-Balkan Ödünçlemeleri; 4AD
Iberia|Gelenek|Neşe|Dans; Spor|Galiçya Pandeireteiras; Kadın Vokal Üçlüsü; Çağdaş Folk
Western Europe; North Africa; Fusion|Elektronik|Karanlık|Yolculuk; Spor|Elektronik; Arap Vokal; Trip-Hop
Latin America|Tel/Enstrümantal; Caz|Neşe|Odak; Spor|Brezilya Gitarı; Bossa/Samba; 60'lar-70'ler
Iberia; Anatolia; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Organic Electronic; Turkish Song; Downtempo
Southern Africa|Caz|Neşe|Dans; Odak; Yolculuk|Piyano Cazı; Güney Afrika Hattı; Cape Town
Western Europe|Vokal/Şarkı; Gelenek|Epik|Dans; Spor|Oksitan Polifoni; Altı Kişilik Vokal; Corrèze
North America|Reggae/Dub|Neşe|Dans; Odak|Cinematic Soul; Enstrümantal Funk; Dub-Soul
Anatolia|Elektronik|Neşe|Dans; Yolculuk; Spor|Elektronik; Anadolu Psychedelia; Organik House
Western Europe|Gelenek; Nefesli/Bando|Melankoli|Gece|Balkan/Klezmer/Fransız Halk Ezgileri; Berlin Kolektifi
Western Europe|Elektronik|Dinginlik|Dans; Odak; Yolculuk|Downtempo; Trip-Hop; Sample Bazlı
Southern Africa; Western Europe; Fusion|Elektronik|Karanlık|Dans; Odak; Spor|Zimbabve Geleneği; Mbira; Elektronik Prodüksiyon
Mediterranean Europe|Psikedelik/Rock|Neşe|Dans; Spor|Sicilya Folk; Arap-Akdeniz Füzyon; Psikedelik Soul
Levant|Klasik/Modern Klasik|Romantik|Odak|Arap Orkestral Pop; Film Müziği; Enstrümantal
Latin America|Klasik/Modern Klasik|Karanlık|Dans|Exotica; And Etkili Vokal; Orkestral Pop
Western Europe|Vokal/Şarkı; Gelenek|Neşe|Dans|Oksitan Kadın Vokal Üçlüsü; Tef; Toulouse
East Africa|Vokal/Şarkı|Melankoli|Dans; Odak; Yolculuk; Spor|Taarab; Zanzibar Geleneği; Svahili Şarkısı
Latin America|Gelenek|Romantik|Dans|Müzik Llanera; Arpa/Cuatro; Joropo
Latin America|Elektronik|Neşe|Dans; Odak; Spor|Meksika Elektronik/Cumbia; Sample Bazlı; Camilo Lara
Western Europe|Gelenek; Nefesli/Bando|Melankoli|Gece|İrlanda Geleneksel Müziği; Tin Whistle; Reel; Jig; Slow Air; Manchester İrlanda Sahnesi
Western Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Gece|British Medieval Music; Renaissance; Early Music; Vocal Polyphony
North America; Western Europe; Fusion|Gelenek; Pop/Şarkı-Yazarı|Neşe|Yolculuk; Spor|Americana; Country; Rock; Bluegrass; İrlanda/Appalachia Kökleri
Western Europe|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk|Ethio-Jazz; Balkan-Doğu Füzyon; Psikedelik Enstrümantal
Mediterranean Europe|Vokal/Şarkı; Gelenek|Neşe|Spor|Akdeniz Vokali; İtalyan Halk Repertuvarı; Milano
North America; Balkans; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Dans; Yolculuk|Balkan Folk; Albanian Song; Violin; Romani and Eastern European Repertoire
Central Africa; Western Europe; Fusion|Gelenek; Elektronik|Melankoli|Dans|Bantu Lullabies; Congolese Percussion; Electronic Hip-Hop
Northern Europe|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Dans; Spor|İsveç Halk Müziği-Arap Pop; Makam (Hicaz-Nihavend, Bayati); Çok Dilli
West Africa; Southern Africa; North America; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece|Mande Vokal; Klasik Gitar Transkripsiyon; Çello
South Asia; Western Europe; Fusion|Reggae/Dub|Neşe|Odak|Sitar-Dub; Sufi Gelenek; Deneysel Prodüksiyon
Western Europe|Hip-Hop/Rap|Neşe|Dans|Fransız Rap; Balkan-Sample Hip-Hop; Kabare
Western Europe; Latin America; Fusion|Vokal/Şarkı; Gelenek|Romantik|Gece; Yolculuk|Ceremonial Folk; Andean Flute; Medicine Song
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Spor|Çağdaş Chanson; İspanyolca-Fransızca
East Africa|Vokal/Şarkı|Epik|Dans|Maloya; Trans Vokal; Ritüel Müzik
Levant; North America|Caz|Neşe|Dans; Odak; Yolculuk|Kontrbas; Ortadoğu Ezgileri + Caz; New York
North Africa|Caz|Melankoli|Dans; Yolculuk|Hassani; Sahra Blues; Caz-Soul Füzyon
Levant; North Africa|Elektronik; Vokal/Şarkı|Melankoli|Dans|Arapça İndie-Elektronik; Beyrut-Kahire Hattı
North America|Vokal/Şarkı; Blues|Neşe|Dans|Erken Caz/Ragtime Canlanması; Western Swing
North Africa|Reggae/Dub|Neşe|Dans; Yolculuk; Spor|Libya Reggae; Ma'luf; Yerel Pop
Western Europe|Caz|Neşe|Dans; Odak|Çello; Caz-Klasik Füzyon; Prodüksiyon
North Africa|Funk/Soul|Melankoli|Dans; Gece; Yolculuk|Arap Funk; Enstrümantal Groove; Retro Prodüksiyon
Iberia|Caz|Melankoli|Odak; Gece|Caz Vokal; Trombon; Bossa Nova-Katalan Folk
North America|Funk/Soul; Psikedelik/Rock|Karanlık|Dans; Odak; Spor|Afro-Soul; Afrobeat; Ethiopian Funk; Proto-Metal
Western Europe|Pop/Şarkı-Yazarı|Romantik|Dans|Fransız Chanson; Edebi Pop; Şarkı Yazımı
Western Europe; Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Dans; Spor|French Chanson; Brazilian Pop; Vocal Duet
Western Europe; Mediterranean Europe|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Downtempo; Trip-Hop; Vintage Sample; Electro Swing; Vintage Pop
Northern Europe|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Yolculuk; Spor|Trip-Hop; Torch Song; Stockholm
Western Europe|Psikedelik/Rock; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Yé-yé Canlanması; Sörf Gitarı; Paris
North America|Tel/Enstrümantal|Neşe|Dans|Old-Time; Afro-Amerikan Yaylı Grup; Banjo Geleneği
Levant; Anatolia; Fusion|Elektronik|Romantik|Yolculuk|Funk; Video-Mashup Prodüksiyon; Elektronik Soul
Western Europe; Caribbean; Fusion|Gelenek; Vokal/Şarkı|Neşe|Odak; Spor|Flemança Calypso; Belçika (Mauro Pawlowski Projesi)
Latin America|Gelenek; Nefesli/Bando|Romantik|Gece|Mixe Women’s Brass Band; Oaxacan Regional Music; Protest Song
North America|Psikedelik/Rock; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Neo-Psychedelia; Folk-Rock; Space Rock; Indie Rock
North America|Psikedelik/Rock|Romantik|Odak; Gece; Yolculuk|Indie Rock; Balkan-Mariachi Etkisi; Sinemasal Folk
Western Europe|Hip-Hop/Rap|Neşe|Dans; Spor|Oksitan Repertuvar; Tamburin-Rap; Sokak Geleneği
Latin America; Western Europe; Fusion|Caz; Gelenek|Melankoli|Yolculuk|Nuevo Tango; Electronica; Chamber Jazz
North America|Elektronik; Vokal/Şarkı|Neşe|Dans; Odak; Spor|New York Global Beat; Turntables on the Hudson
North America|Reggae/Dub|Karanlık|Dans; Odak; Spor|Cinematic Soul; Enstrümantal Funk; Dub-Soul
Western Europe|Pop/Şarkı-Yazarı; Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|Avangart Şanson; Tiyatro; Areski İşbirliği
Oceania|Vokal/Şarkı; Blues|Romantik|Dans; Gece|Country-Soul; Rockabilly; Güçlü Vokal
Latin America|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk|Nu-Cumbia; Dijital Folklor; Chamamé
Western Europe|Elektronik|Neşe|Dans; Spor|Indie Pop; Deneysel Folk; Elektronik
Western Europe|Funk/Soul|Dinginlik|Yolculuk|Nyah-Rock; Britanya Funk'ı; Karayip Diasporası
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Dans; Spor|İngiliz Folk Geleneği; Carthy Hattı
Western Europe; Latin America; Fusion|Elektronik|Neşe|Odak; Yolculuk|Kolombiya Pasifik Geleneği; Dub; Currulao
Latin America|Elektronik|Karanlık|Dans; Yolculuk|Deneysel Elektronik; Other People; New York-Santiago
Western Europe|Caz; Elektronik|Neşe|Dans; Spor|Ambient Jazz; Electronic Jazz
Western Europe|Caz|Neşe|Gece; Yolculuk|Folk-Hop; Swing; Sokak Cazı
Latin America|Pop/Şarkı-Yazarı; Caz|Neşe|Dans; Spor|Bossa Nova + Amerikan Pop; Brasil '66
Latin America|Pop/Şarkı-Yazarı|Romantik|Dans|Brezilya Şarkı-Yazarlığı; Los Hermanos; Film Müziği
North America|Pop/Şarkı-Yazarı|Melankoli|Dans; Odak; Yolculuk|Amerikan Indie Folk; Anlatıcı Söz
Western Europe|Psikedelik/Rock|Neşe|Dans|Spoken Word; Post-Punk
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Anti-Folk/İndie; İngilizce Söz; Paris
Anatolia|Psikedelik/Rock|Neşe|Yolculuk|Arabesk Rock; Psychedelic Rock; Turkish Alternative Rock
Western Europe|Pop/Şarkı-Yazarı; Psikedelik/Rock|Neşe|Dans|Art Pop; Indie Rock
Western Europe|Elektronik|Romantik|Yolculuk|Trip-Hop; Lounge; Downtempo
Caribbean|Reggae/Dub; Caz|Neşe|Odak; Yolculuk|Ska/Reggae Gitarı; Stüdyo Sesyonları; Studio One
Latin America; North America; Fusion|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Bossa Nova; Samba; Brazilian Jazz
Latin America|Elektronik|Neşe|Dans; Spor|Cumbia; Dub; Tropical Electronic
East & Central Asia|Caz; Gelenek|Neşe|Yolculuk|Shakuhachi + Caz; 'Bamboo' (1970)
Western Europe|Caz; Pop/Şarkı-Yazarı|Neşe|Dans|Londra Sahnesi; Afro-Latin Etkili İndie
Levant; West Africa; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece|Piyano-Gitar Doğaçlaması; İsrail-Mali Füzyonu
Anatolia|Psikedelik/Rock|Karanlık|Dans; Yolculuk; Spor|Anadolu Psych; Sürrealist Sahne Kimliği; İstanbul
Levant|Psikedelik/Rock|Neşe|Yolculuk|Psychedelic Rock; Akdeniz Surf; Doğu Rock
North America; Western Europe; North Africa; Fusion|Elektronik; Psikedelik/Rock|Karanlık|Odak; Spor|Psychedelic Dance; Arabic Rhythms; Samples/Synths
Anatolia|Psikedelik/Rock; Funk/Soul|Neşe|Dans|Anadolu Psych Derlemesi/Edit; Freak-Funk; Disco-Punk; Fat City / B-Music
Northern Europe|Psikedelik/Rock|Karanlık|Dans; Yolculuk; Spor|Psikedelik Rock; Afro-Krautrock; Maskeli Kolektif
East & Central Asia|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk|Japon Psych; Sitar + Gitar; Guruguru Brain
Anatolia|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Odak; Spor|Anadolu Funk; Stüdyo Düzenlemeleri; 1973
Southeast Asia|Psikedelik/Rock|Neşe|Yolculuk|Kamboçya Psychedelia; Surf Rock; Garage
North America|Psikedelik/Rock; Blues|Karanlık|Yolculuk|Ağır Psych / Hard Rock; Füz Gitar + Mouth Harp; 1972 Iowa Bodrum Kaydı
Western Europe|Psikedelik/Rock|Karanlık|Dans; Yolculuk; Spor|Retro Hard Rock; Flüt-Rock; Blues-Rock
Latin America|Elektronik|Neşe|Dans; Odak; Yolculuk|Deneysel Elektronik; Other People; New York-Santiago
North America; Latin America; West Africa; Fusion|Funk/Soul; Psikedelik/Rock|Neşe|Odak; Spor|Enstrümantal Gitar; Chicha/Tuareg/Son Montuno; New Orleans
Northern Europe|Caz; Funk/Soul|Neşe|Dans; Odak; Yolculuk|Afrobeat; Jazz-Funk
Central Africa|Blues|Melankoli|Gece|Bassa Folk; Minimal Akustik; Afro-Blues
North America|Pop/Şarkı-Yazarı|Melankoli|Gece|Québec Chanson; Karkwa; Fransızca Söz
Western Europe; Levant|Pop/Şarkı-Yazarı|Melankoli|Gece|French Chanson; Indie Pop; Vocal Duet
Iberia; Western Europe|Vokal/Şarkı; Gelenek|Neşe|Dans|Mediterranean Song; Piano; Multi-Instrumental Folk
North America; Latin America|Psikedelik/Rock|Dinginlik|Odak; Gece; Yolculuk|Peru Chicha'sı; Cumbia + Sörf Gitar; Brooklyn
Caribbean|Gelenek; Vokal/Şarkı|Neşe|Gece|Antiller Biguine/Gitar; Disques Debs Kataloğu
West Africa|Blues|Epik|Yolculuk|Mali Blues; Kayes Gitar; Solo Vokal-Gitar
Latin America|Gelenek; Psikedelik/Rock|Neşe|Dans; Odak|Amazon Cumbia'sı; Iquitos; Kendi Etiketi
East & Central Asia|Caz; Vokal/Şarkı|Epik|Dans|Caz Vokali; Geniş Teknik Aralık; Paris-Seul
North Africa; West Africa|Psikedelik/Rock; Gelenek|Neşe|Dans; Yolculuk; Spor|Tuareg Assouf; Çöl Blues; Tamasheq; Sahra Gitarı
Caucasus; Western Europe; Anatolia; Fusion|Hip-Hop/Rap|Neşe|Yolculuk|Ermeni Halk Müziği; Anadolu Füzyon; Hip-Hop
Balkans|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak; Spor|Makedon/Balkan Repertuvar; Çoklu Enstrüman
North America; Latin America|Vokal/Şarkı|Neşe|Yolculuk|Latin Amerika Repertuvarı; Kadın Vokal İkilisi; Brooklyn
Western Europe|Psikedelik/Rock; Elektronik|Neşe|Yolculuk; Spor|Kozmik Psych/Krautrock; Viyana
Western Europe; North America|Funk/Soul; Hip-Hop/Rap|Neşe|Dans|Cinematic Soul; Funk; Guest Vocal
East Africa|Funk/Soul; Gelenek|Neşe|Yolculuk|Somali Funk; Reggae Groove; 1970'ler-80'ler Mogadişu; Devlet Destekli Topluluk; Çağrı-Yanıt
West Africa|Vokal/Şarkı|Melankoli|Yolculuk|Togo Vokali; Ewe/Fransızca Söz; 70'ler
Northern Europe|Caz|Neşe|Dans; Odak; Spor|Enstrümantal; Kütüphane Müziği; Anadolu-Japon Etkili Caz-Funk
Western Europe; North America; West Africa; Latin America; Fusion|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk; Spor|West African Griot Song; Jazz Guitar; Ngoni; Senegalese Rhythm
Western Europe; East Africa; Fusion|Caz|Neşe|Dans; Odak|Ethio-Jazz; Enstrümantal Funk
Iberia; West Africa; Fusion|Tel/Enstrümantal; Gelenek|Romantik|Odak; Gece|Flamenko + Kora; 'Songhai' Hattı
Latin America; Western Europe; Fusion|Vokal/Şarkı|Dinginlik|Gece|Kamera Folk; Çok Dilli Vokal Düeti
West Africa; North America; Fusion|Gelenek; Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Mande Griot; Balafon/Ngoni; String Quartet
North Africa|Elektronik|Melankoli|Dans; Gece; Yolculuk|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Anatolia; Iberia; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Flamenco Guitar; Turkish Clarinet; Instrumental Duet
Western Europe|Tel/Enstrümantal|Neşe|Dans; Odak; Gece|Enstrümantal; Doğu Avrupa Füzyon; Sinemasal Folk
Anatolia|Vokal/Şarkı; Gelenek|Melankoli|Dans|Turkish Folk; Black Sea Music; Turkish Classical Song; Tango; Arabesk
Western Europe|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans; Spor|Fransız Sanat-Pop; Agnès Gayraud; Paris
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans|Folk-Blues; Akdeniz/Mande İşbirlikleri; Cévennes
Southeast Asia|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk|Kamboçya Psychedelia; Surf Rock; Garage
West Africa|Gelenek|Romantik|Dans; Yolculuk|Kora; Mande Geleneği; Griot Müziği
Levant; East & Central Asia; Fusion|Funk/Soul|Neşe|Yolculuk|Ortadoğu Funk; Enstrümantal Groove; Retro Prodüksiyon
Levant; Western Europe|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Dans; Gece; Yolculuk|Çoklu Enstrüman; Sokak İcrası; Loop
Northern Europe; Balkans; Fusion|Caz; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Contemporary Jazz; Balkan Modal Jazz
Levant; Iberia; Fusion|Caz; Gelenek|Romantik|Dans; Spor|Oud; Flamenco; Jazz Improvisation
Levant; Anatolia|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Oud; Middle Eastern Percussion; Ottoman Music; Jewish Musical Traditions
West Africa|Funk/Soul; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Afrobeat; Kuti Ailesi'nin Üçüncü Kuşağı; Lagos
Western Europe; Caribbean|Elektronik; Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|Indie Dance; Remix; Steelpan Artist
Caribbean; North America|Vokal/Şarkı; Caz|Neşe|Yolculuk|Afro-Kuba + Caz; Kadın İkili; Toronto
Western Europe|Gelenek; Vokal/Şarkı|Epik|Gece; Yolculuk|İngiliz Halk Müziği; Dobro/Slide; Keman
Mediterranean Europe; Balkans; North Africa; Fusion|Gelenek; Nefesli/Bando|Epik|Yolculuk|Salento; Balkan/Akdeniz Repertuvarı; Saz-Klarnet-Akordeon
Balkans; Eastern Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Keman/Citera; Vojvodina Macar Hattı
Eastern Europe|Vokal/Şarkı; Gelenek|Neşe|Gece|Macar Roman (Lovari) Vokal Geleneği
Iranian World|Elektronik|Neşe|Yolculuk|Fars Pop; Elektronik; Diaspora Funk
North Africa|Caz; Nefesli/Bando|Neşe|Dans; Yolculuk|Egyptian Jazz; Free Jazz; Big Band
Western Europe; Southern Africa; Fusion|Vokal/Şarkı; Elektronik|Melankoli|Gece|World Pop; Afrobeat; Experimental Song
East Africa|Psikedelik/Rock|Epik|Dans; Yolculuk; Spor|Maloya-Rock; Trans Perküsyon; Ritüel Müzik
West Africa|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk; Spor|Highlife; Hiplife; Afrobeat
Western Europe; Anatolia; Iberia; Fusion|Tel/Enstrümantal; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Beş Telli Kontrbas; Kemençe; Doğu Makamları; Akdeniz; Endülüs Cante Jondo
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans|Fransız Alternatif Şanson; Sade Gitar; Bariton
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Çağdaş Chanson; Folk; Paris
Southern Africa|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Semba; Angola Repertuvarı; Sürgün Şarkıları
Western Europe; Central Africa|Vokal/Şarkı; Hip-Hop/Rap|Melankoli|Dans|Kongo Vokali + Fransız Rapı; Diaspora
Iberia|Tel/Enstrümantal|Melankoli|Gece|Portekiz Noir Gitarı; Enstrümantal İkili; Lizbon
Iberia|Klasik/Modern Klasik|Dinginlik|Dans; Odak; Gece|Minimalist Kamara; Film Müziği; Portekiz Yeni Klasik
Southeast Asia|Funk/Soul|Neşe|Dans; Spor|Vietnam Soul; 1960'lar Saigon Pop; Psikedelik Funk
South Asia|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Yolculuk|Urduca Şiir + Minimal Düzenleme; Gazel Esini
North Africa|Klasik/Modern Klasik; Caz|Epik|Spor|Film Müziği; Kütüphane Müziği; Habibi Funk
North America|Elektronik|Karanlık|Dans; Odak; Spor|Downtempo/Dub; Bossa-Reggae Ödünçlemeleri; Washington DC
West Africa; North America; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Yolculuk|Manding Kora; Ngoni; Balafon; Jazz Bass
West Africa|Vokal/Şarkı|Dinginlik|Yolculuk|Bamana Repertuvarı; Ngoni + Elektrikli Gitar; Bağımsız Söz
West Africa|Reggae/Dub|Karanlık|Dans|Afrika Reggae'si; Politik Söz; Sürgün
West Africa; Western Europe|Pop/Şarkı-Yazarı; Funk/Soul|Melankoli|Dans|Mali Pop; Bambara/Fransızca Söz; Aktivizm
Levant|Caz; Elektronik|Neşe|Odak|İsrail Caz Sahnesi; Davul/Prodüksiyon; Tel Aviv
Latin America|Gelenek; Funk/Soul|Neşe|Dans; Odak|Kolombiya Arşivi + Çağdaş Kadro; Discos Fuentes Stüdyosu
Southern Africa|Vokal/Şarkı; Gelenek|Melankoli|Dans; Gece|Mbira; Shona Geleneği; Zimbabve
Caribbean|Hip-Hop/Rap; Funk/Soul|Neşe|Dans; Spor|Flüt + Rap; Cumbia/Dancehall; Paris
Balkans|Elektronik; Gelenek|Romantik|Dans; Odak|Romen Lounge/Pop Canlanması; 60'lar-70'ler Arşivi
Western Europe|Pop/Şarkı-Yazarı; Psikedelik/Rock|Romantik|Spor|İndie Rock + Fransız Pop Düeti; Scotland
Western Europe|Caz; Psikedelik/Rock|Karanlık|Odak; Yolculuk|Jazz-Funk; Psychedelic; Library Music; Experimental
North America|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Spor|Psychedelic Soul; Cajun/Louisiana; Özel Basım 45'lik; 1972
Latin America|Gelenek|Epik|Dans; Spor|Umbanda/Candomblé Ritüel Repertuvarı; Perküsyon-Vokal
West Africa|Funk/Soul; Gelenek|Neşe|Odak; Yolculuk|Guinea Ulusal Orkestrası; Afro-Küba; Mande Ezgileri
Iberia|Elektronik|Karanlık|Gece; Yolculuk|Deneysel Flamenko; Performans; Ses Sanatı
West Africa; Western Europe|Vokal/Şarkı; Elektronik|Melankoli|Dans|Bambara Vokali + Avrupa Prodüksiyonu; Paris
Western Europe; Latin America; Fusion|Elektronik; Gelenek|Melankoli|Odak; Gece; Yolculuk|Organic House; Electrocumbia; Acoustic Ensemble
West Africa|Gelenek|Epik|Dans; Yolculuk; Spor|Yoruba Davul Geleneği; Djembe/Konga; ABD Yerleşik
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Balkan Elektro; Bucovina Club; Frankfurt
Western Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Romantik|Dans|Fransız Pop; Cezayir Kökeni; Oyuncu-Şarkıcı
Western Europe|Gelenek|Romantik|Dans|Indie Folk; Chanson; Karanlık Pop
Anatolia|Psikedelik/Rock|Neşe|Yolculuk|Anadolu Rock'ın Kurucuları; 1967; Bağlama + Elektrikli Grup
Northern Europe|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Dans; Spor|İsveç Halk Müziği-Arap Pop; Makam (Hicaz-Nihavend, Bayati); Çok Dilli
Central Africa; Southern Africa; Western Europe|Funk/Soul; Vokal/Şarkı|Neşe|Dans; Spor|Afro-Funk; Swahili Pop; African Disco; Vocal Group
North America|Psikedelik/Rock; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|Art-Rock; Talking Heads; Luaka Bop Kurucusu
North America|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|İndie Folk-Rock; Wainwright Ailesi; Montreal
Southern Africa|Gelenek; Klasik/Modern Klasik|Epik|Dans; Spor|Mbira; Traditional Zimbabwean Song
West Africa|Vokal/Şarkı|Neşe|Dans; Spor|Funaná/Batuque; Kreol Söz; Yeni Kuşak
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans; Gece|İngiliz Folk; Akustik Düet; Gitar-Vokal
Iberia|Nefesli/Bando|Neşe|Yolculuk; Spor|Balkan Brass; Roman Müziği; Klezmer Füzyon
Northern Europe; West Africa; Fusion|Blues; Tel/Enstrümantal|Neşe|Dans; Odak; Gece; Yolculuk|Afro-Blues; Gitar-Perküsyon Düeti; Batı Afrika Ritmi
Anatolia; Mediterranean Europe; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Anadolu-Akdeniz Hattı; Çok Uluslu Kadro
Iranian World; Latin America; Fusion|Gelenek|Dinginlik|Odak; Gece|Fars–And Füzyonu; Kamançe/Charango-Quena
Balkans|Gelenek; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Kanun + Ud + Ney; Yunan Modal Topluluğu
Mediterranean Europe|Caz; Tel/Enstrümantal|Neşe|Dans; Odak|Mediterranean Jazz; Acoustic Instrumental
Western Europe|Vokal/Şarkı|Melankoli|Gece|Çok Dilli Folk; Kamera Pop; Vokal Üçlüsü
Western Europe; Latin America; East Africa; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Dans|Brazilian Song; Breton Folk; Maloya; Vocal Harmony; Percussion
East & Central Asia|Gelenek; Psikedelik/Rock|Dinginlik|Odak; Yolculuk|Taiwanese Indie; Chinese Folk; Rock
Caucasus; Anatolia; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Duduk + Kemençe; Making Tracks Rezidansı
Anatolia|Elektronik|Dinginlik|Odak; Gece; Yolculuk|Elektronik; Anadolu Psychedelia; Organik House
Caucasus|Caz|Melankoli|Odak; Yolculuk|Mugam-Caz Piyanosu; Bakü
Latin America|Gelenek|Neşe|Dans; Spor|Carimbó; Pará/Amazon; 'Carimbó Kralı'
Western Europe|Pop/Şarkı-Yazarı; Klasik/Modern Klasik|Epik|Dans; Yolculuk; Spor|Piyano-Vokal; Neo-Klasik Soul; Londra
Iberia; Anatolia; Fusion|Gelenek|Melankoli|Gece; Yolculuk|Klezmer; Balkan Füzyon; Sefarad Repertuvarı
North America|Pop/Şarkı-Yazarı|Melankoli|Gece|Indie Folk-Pop; Broken Social Scene Çevresi
Levant|Pop/Şarkı-Yazarı; Blues|Neşe|Dans; Spor|Folk-Rock; İngilizce Söz; Tel Aviv
Western Europe|Psikedelik/Rock; Elektronik|Neşe|Dans; Yolculuk; Spor|Art-Rock; Vokal Harmoni + Elektronik; Londra
North America|Funk/Soul|Romantik|Dans|R&B/Soul Canlanması; Analog Kayıt; Kaliforniya
Eastern Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Odak; Yolculuk|Klezmer; Yahudi Kamara Müziği; Doğaçlama
Western Europe; South Asia; Fusion|Elektronik; Gelenek|Neşe|Dans; Yolculuk; Spor|Tropical Electronic; Cumbia; Indian Instrumental
Levant|Reggae/Dub|Neşe|Dans; Spor|Reggae; Çöl Blues; Arap Rock
Western Europe|Gelenek; Elektronik|Karanlık|Odak; Yolculuk|Breton Dañs Plinn; Brass; Electro-Folk
Anatolia|Elektronik|Karanlık|Dans; Odak; Yolculuk; Gece|İstanbul Downtempo; Türk Ezgileri + Elektronik
Western Europe|Elektronik; Funk/Soul|Neşe|Dans; Odak; Spor|Nu-Disko/Funk Elektroniği; Leiden
Latin America|Tel/Enstrümantal|Epik|Odak|Akustik Gitar Düeti; Flamenko-Metal Füzyon
Western Europe; Iberia; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans; Odak|Sefarad/Ladino Repertuvarı; Londra
West Africa; Western Europe; Fusion|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Kora-Klarnet-Çello; Kamara Füzyon
Eastern Europe; Northern Europe; Fusion|Gelenek|Melankoli|Dans; Yolculuk; Spor|Dark Folk; Alternatif Folk; Polonya ve İskandinav Gelenekleri; Arkaik Çalgılar; Trans Dansı
Balkans; Western Europe; Fusion|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Girit Lavtası + Elektrikli Arp; Girit–Londra
Mediterranean Europe|Klasik/Modern Klasik|Neşe|Dans|Napoli Tarantella; Sokak Orkestrası; Tammurriata
Mediterranean Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak|Chitarra Battente; Southern Italian Folk; Guitar Duo
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|İngiliz Geleneksel Şarkı; Birmingham; Arşiv Araştırması
Western Europe; West Africa; Fusion|Reggae/Dub|Epik|Dans; Yolculuk|Fransız Reggae; Ragga; Afro-Reggae
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Folk-Blues; Akdeniz/Mande İşbirlikleri; Cévennes
Northern Europe; Levant; Fusion|Caz|Dinginlik|Odak; Gece|Caz Piyanosu; Maqam Çalışmaları; Mikrotonal Klavye
Western Europe; North Africa|Hip-Hop/Rap; Vokal/Şarkı|Neşe|Dans|Fransız-Cezayir Rap/Şanson; Sahne Kimliği; Lyon
Latin America; Western Europe; Fusion|Vokal/Şarkı|Neşe|Gece|Kamera Folk; Çok Dilli Vokal Düeti
North America|Gelenek; Blues|Melankoli|Gece|Siyah Telli Çalgı Geleneği; Banjo/Keman; Pulitzer
Western Europe; Balkans; Fusion|Caz; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Jazz Manouche + Balkan/Klezmer
Mediterranean Europe|Pop/Şarkı-Yazarı|Neşe|Dans|İtalyan Cantautore; Edebi Rock; Akdeniz Anlatısı
Balkans|Vokal/Şarkı; Gelenek|Melankoli|Gece|Bulgar Kadın Polifonisi; Pirin Bölgesi
North Africa|Vokal/Şarkı|Melankoli|Dans; Yolculuk; Spor|Chaâbi; Cezayir Şarkı Geleneği; Mandol
North America|Gelenek|Romantik|Dans; Yolculuk|Québec Bluegrass; Fransızca Söz; Mizah
North America; Levant; Western Europe|Klasik/Modern Klasik; Vokal/Şarkı|Neşe|Dans|Venetian Jewish Song; Tenor; Chamber Ensemble
Iberia|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak|Portuguese Folk; Mandolin; Cavaquinho; Contemporary Acoustic
Balkans|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Aegean Folk; Ikariotikos; Kanun/Oud/Lyra
Eastern Europe|Klasik/Modern Klasik|Melankoli|Gece|Cello Duo; Improvisation; East European Folk
Western Europe|Hip-Hop/Rap; Elektronik|Neşe|Dans; Spor|Trip-Hop; Instrumental Hip-Hop; Sampling; Guest Vocal
North Africa|Vokal/Şarkı|Neşe|Yolculuk|Mısır Şarkı Geleneği; Kahire; 60'lar-70'ler
West Africa|Psikedelik/Rock|Karanlık|Yolculuk; Spor|Tuareg Rock; Çöl Blues; Psikedelik
Caucasus; North America|Caz; Gelenek|Neşe|Dans; Yolculuk|Armenian Jazz; Percussion; Piano Improvisation
Western Europe|Gelenek; Nefesli/Bando|Neşe|Yolculuk|Klezmer + Balkan; Sokak Bandosu; Amsterdam
Western Europe; West Africa; Fusion|Reggae/Dub|Neşe|Dans; Spor|Fransız Reggae; Ragga; Afro-Reggae
Western Europe; Latin America; Fusion|Elektronik; Gelenek|Neşe|Yolculuk|Kolombiya Pasifik Geleneği; Dub; Currulao; Deep Funk / Tru Thoughts Hattı
Levant|Tel/Enstrümantal|Neşe|Dans; Odak; Yolculuk|Ud/Gitar; Arap Repertuvarı
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Experimental Jazz; Contemporary Jazz
Anatolia; Iranian World; Fusion|Tel/Enstrümantal; Caz|Melankoli|Odak; Gece; Yolculuk|Fretless Guitar; Iranian Tanbur/Setar; Modal Improvisation
Latin America|Gelenek|Neşe|Dans; Spor|Carimbó; Amazon Pop; Lambada
West Africa; Western Europe; Fusion|Psikedelik/Rock; Blues|Epik|Yolculuk|Tuareg Rock; Ağır Gitar; Brüksel
Northern Europe|Gelenek|Epik|Gece; Yolculuk|İskandinav Ritüel Müziği; Tarihsel Çalgılar; Nordic Folk
Eastern Europe|Reggae/Dub|Neşe|Spor|Rus Halk Müziği; Folk-Rock; Ska Etkisi
Latin America|Vokal/Şarkı; Reggae/Dub|Melankoli|Gece|MPB; Afro-Brezilya Ritimleri; Pop; Caz; Dub
Latin America|Vokal/Şarkı; Gelenek|Melankoli|Dans|Tango Canción; Buenos Aires; Edebi Söz
North Africa|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Tuareg Rock; Çöl Blues; Assouf
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Yolculuk; Spor|Çağdaş Chanson; İspanyolca-Fransızca
Iberia|Gelenek; Vokal/Şarkı|Melankoli|Gece|Galiçya Halk Müziği; Galiçyaca Söz
Caribbean; Western Europe; Fusion|Reggae/Dub; Funk/Soul|Romantik|Dans; Spor|Roots Reggae; Accordion; Afrobeat; Maloya; Soul
Caribbean|Reggae/Dub; Vokal/Şarkı|Neşe|Gece|Ska / Rocksteady; Erken Reggae
Western Europe; South Asia; Levant; Fusion|Caz; Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Politik Dünya Müziği Kolektifi; Al MacSween/Modarelli
Western Europe; Latin America|Vokal/Şarkı; Elektronik|Dinginlik|Odak; Gece; Yolculuk|Kolombiya Ezgileri + Avrupa Prodüksiyonu; Brüksel
Levant; North America; Balkans; Fusion|Elektronik; Nefesli/Bando|Neşe|Dans; Odak; Spor|Balkan Brass; Mediterranean; Hip-Hop; Dub
Mediterranean Europe; North Africa; Latin America; Caribbean; Fusion|Caz; Funk/Soul|Neşe|Dans; Odak; Yolculuk; Spor|Hyperjazz; Afrobeat; Neapolitan Jazz; Improvisation; Mediterranean Trance
North America|Pop/Şarkı-Yazarı; Funk/Soul|Neşe|Dans; Spor|Detroit Folk-Soul; 'Cold Fact'; Sugar Man
Latin America|Elektronik|Karanlık|Dans|Sample Kolajı; 'Dijital Macumba'; Psikedelik Edit; Sähkö Recordings
Levant; Iberia; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak|Oud; Flute/Saxophone; Flamenco Jazz
Eastern Europe|Klasik/Modern Klasik; Elektronik|Dinginlik|Odak; Gece; Yolculuk|Piyano; Neo-Klasik/Ambient; Varşova
Latin America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Bossa Nova; MPB; Samba
Anatolia|Caz; Tel/Enstrümantal|Neşe|Dans; Odak; Yolculuk; Spor|Jazz Fusion; Fretless Guitar; Piano; Percussion
Caribbean; Western Europe|Hip-Hop/Rap; Caz|Melankoli|Dans; Yolculuk|Sözlü Şiir + Caz/Funk; Trinidad-Londra Hattı
North America|Caz; Vokal/Şarkı|Dinginlik|Dans; Gece; Yolculuk|Vokal Doğaçlama; Çok Oktavlı Teknik
North America|Caz; Vokal/Şarkı|Melankoli|Gece|Caz Vokali; Blues/Şanson Etkisi
Mediterranean Europe|Gelenek|Epik|Yolculuk; Spor|Pizzica; Tarantella; Salento Halk Müziği
Iberia|Caz; Vokal/Şarkı|Melankoli|Gece|Klarnet + Vokal; Katalan Şiiri; Barselona
East & Central Asia|Gelenek; Psikedelik/Rock|Neşe|Yolculuk|Tuvan Throat Singing; Kargyraa; Folk-Rock; Ethnic Electronica
Iberia|Caz; Gelenek|Melankoli|Dans; Yolculuk|Bambu Flüt/Saksofon; Fado + Doğu Ezgileri; Lizbon
Western Europe; West Africa; Fusion|Elektronik; Gelenek|Neşe|Dans; Odak; Yolculuk|Gnawa; Electronic; West African Roots
Iberia|Funk/Soul|Neşe|Dans; Spor|60'lar Soul/R&B Canlanması; Barselona
North Africa|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Elektrikli Oud; Malouf; Rock Füzyon
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Yolculuk|Meksika Şarkısı; Latin Amerika Halk/Popüler Şarkısı; Son Jarocho/Bolero Bağlamı; Rap/Spoken-Word Katkısı
Levant|Caz|Neşe|Dans; Odak; Yolculuk; Spor|Saksofon; Arap Ezgileri + Caz; Paris
Western Europe|Klasik/Modern Klasik|Dinginlik|Dans; Odak|Kontrabas; Erken Müzik; Doğu-Batı Kamara Füzyonu
Western Europe|Tel/Enstrümantal|Neşe|Dans; Odak|Enstrümantal; Doğu Avrupa Füzyon; Sinemasal Folk
Levant|Caz|Dinginlik|Odak; Yolculuk|Mandolin; Ortadoğu Cazı; Kamara Füzyon
Latin America|Vokal/Şarkı; Elektronik|Dinginlik|Dans|Latin Amerika Vokalleri + Afro-Peru Elektroniği
North America; Latin America|Gelenek; Pop/Şarkı-Yazarı|Dinginlik|Dans|Meksika Son Jarocho + Americana; Jarana/Quijada
Latin America|Psikedelik/Rock; Gelenek|Neşe|Odak; Spor|Psikedelik Latin Soul; Vintage Cumbia; Avándaro Estetiği
Iberia|Elektronik|Neşe|Dans; Yolculuk; Spor|Galiçya Folk; Elektronik; Pandeireteiras
Balkans|Vokal/Şarkı; Nefesli/Bando|Neşe|Dans; Spor|Makedon Roman Repertuvarı; Çalgi/Nefesli
Balkans|Nefesli/Bando|Neşe|Dans; Odak; Yolculuk; Spor|Bulgar Roman Düğün Müziği; Saksafon; Wedding Band
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Alternatif Folk; İngiliz Halk Müziği; Gypsy Jazz; Old-Time
North America|Tel/Enstrümantal; Gelenek|Melankoli|Odak; Gece; Yolculuk|Kalimba; Flute; Djembe; Percussion Duo; World Instrumental
Iranian World|Klasik/Modern Klasik|Melankoli|Yolculuk|Klasik İran Vokali; Kamara Orkestra; Çok Kültürlü Beste
Western Europe; Fusion|Gelenek; Klasik/Modern Klasik|Dinginlik|Odak; Gece; Yolculuk|Kültürlerarası Beste; Küresel Geleneksel Çalgılar; Solo Çoklu Enstrümantalizm
Western Europe; Mediterranean Europe; Fusion|Gelenek; Nefesli/Bando|Neşe|Dans|Güney İtalya Tarantella/Pizzica + Fransız Sokak Topluluğu
North America|Gelenek|Dinginlik|Dans; Gece|Appalachia Vokal; Aktivist Folk; Banjo-Fiddle
Balkans; Anatolia|Gelenek; Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk; Spor|Kıbrıs Halk Müziği; Tzouras/Trombon; Lefşkoşa
East & Central Asia|Gelenek|Epik|Dans|Altay Etno; Komüz/Topşuur, Tılıkı-Şuur, Boğaz Şarkıcılığı
Eastern Europe; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk; Spor|Hungarian-Romani Violin; Polish Carpathian Strings; Improvisation
East & Central Asia|Gelenek; Vokal/Şarkı|Neşe|Yolculuk; Spor|Buryat Repertuvarı; Yokhor Çember Dansı; Gırtlak Vokali
Western Europe; West Africa; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Kelt Arpı + Kora İkilisi
Iranian World|Klasik/Modern Klasik|Melankoli|Dans; Gece; Yolculuk|Fars Pop; Klasik İran Vokali; Los Angeles Diaspora
Anatolia|Psikedelik/Rock|Neşe|Dans|Anadolu Rock'ın Kurucuları; 1967; Bağlama + Elektrikli Grup
Iberia|Elektronik|Karanlık|Odak|Endulüs Elektronik; Flamenko-Techno; Sevillanas
East & Central Asia|Psikedelik/Rock|Karanlık|Odak; Spor|Japon Eleki; Surf-Punk; Enstrümantal Kaos
Eastern Europe|Gelenek|Romantik|Yolculuk|Polonya Köy Repertuvarı; Beyaz Ses; Hurdy-Gurdy/Yaylılar
North Africa|Caz|Dinginlik|Odak; Gece; Yolculuk|Oud; Arap Klasik Müziği; Kamara Caz
Iberia|Nefesli/Bando|Epik|Dans; Yolculuk|Balkan Brass; Roman Müziği; Klezmer Füzyon
Latin America|Gelenek|Neşe|Dans|Amazon Ritimleri; Carimbó; Kuzey Brezilya Folk
Anatolia; Balkans; Iranian World; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Modal Müzik Topluluğu; Kopuz-Lira-Ney-Zarb; Ross Daly / Labyrinth Çevresi
Southern Africa|Psikedelik/Rock|Neşe|Yolculuk|Deneysel Rock; Afro-Psych; Johannesburg
West Africa; Northern Europe|Gelenek; Tel/Enstrümantal|Neşe|Dans; Spor|Kora; Griot Ailesi; İsveç İşbirlikleri
Caribbean|Reggae/Dub|Neşe|Dans; Yolculuk|Roots Reggae; Dub; Kingston
Caribbean|Gelenek|Neşe|Dans|Kuba Geleneği; Afro-Karayip Perküsyonu
Eastern Europe|Gelenek; Caz|Melankoli|Dans|Ukrayna-Polonya Vokali + Caz Kadrosu
Western Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Odak; Yolculuk|Ambient Folk; Scottish Instrumentation; Neo-Classical; Hypnotic Loops
Western Europe|Gelenek|Dinginlik|Odak; Gece; Yolculuk|Balfolk; Melodeon + Gitar; Breton–Gavot–Mazurka Dans Formları
South Asia|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk|Baul; Bengali Devotional; Ektara/Tabla
East & Central Asia|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk; Spor|Buryat Repertuvarı; Yokhor Çember Dansı; Gırtlak Vokali
Mediterranean Europe|Reggae/Dub|Neşe|Odak|Tarantella Dub; Salento Folk; Elektronik
Central Africa; Southern Africa; Fusion|Elektronik; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Tribal House; African Vocal
Northern Europe; West Africa; Caribbean; East & Central Asia; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece; Yolculuk|Balafon/Kora; Batı Afrika Repertuvarı; Kültürlerarası Topluluk
Northern Europe; Iranian World; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Persian Poetry; Global Strings; Nyatiti; Ngoni; Sitar; World Ensemble
Eastern Europe|Tel/Enstrümantal|Melankoli|Gece|Çok Dilli Polonya Repertuvarı; Azınlık Dilleri; Akordeon
Caucasus|Vokal/Şarkı|Neşe|Dans; Spor|Gürcü Halk Müziği; Panduri; Vokal Üçlüsü
North America|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Enstrümantal Gitar; Amerikan Primitif; Ambient Folk
Iberia|Klasik/Modern Klasik|Melankoli|Gece|Polifonik Vokal Düeti; Akdeniz İlahileri; Erken Müzik; Catalonia
Balkans|Vokal/Şarkı|Melankoli|Gece|Akdeniz Vokal Geleneği; Serbest Doğaçlama; Erken Müzik
Latin America|Gelenek; Nefesli/Bando|Melankoli|Yolculuk|Chamamé; Akordeon Geleneği
West Africa|Vokal/Şarkı; Funk/Soul|Neşe|Dans|Benin Vokali; Afro-Pop; Uluslararası İşbirlikleri
North America|Vokal/Şarkı; Gelenek|Melankoli|Dans; Spor|Old-Time/Country; Batı Virginia; Sokak Müzisyenliğinden
Iranian World|Psikedelik/Rock|Melankoli|Yolculuk|İran Psikedelisi; 70'ler Tahran; Fuzz Gitar
North Africa; Western Europe; Fusion|Gelenek; Elektronik|Melankoli|Odak; Gece; Yolculuk|Tunisian Stambeli; Quarter-Tone Violin; Experimental Electronics
North America|Funk/Soul|Neşe|Dans; Odak; Spor|Enstrümantal Soul/Afrobeat; Brooklyn
North America; Balkans|Pop/Şarkı-Yazarı; Nefesli/Bando|Neşe|Dans; Odak|Balkan Nefesli Etkili İndie; Zach Condon
East & Central Asia|Caz; Vokal/Şarkı|Romantik|Gece|Caz Vokali; Geniş Teknik Aralık; Paris-Seul
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans|Folk-Blues; Akdeniz/Mande İşbirlikleri; Cévennes
Western Europe|Caz; Pop/Şarkı-Yazarı|Dinginlik|Dans; Odak; Yolculuk; Spor|Fransız Cazı + Şarkı-Yazarı Vokali; Saksofon
Western Europe|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Çok Dilli Folk; Kamera Pop; Vokal Üçlüsü
Balkans|Psikedelik/Rock|Melankoli|Gece|Yunan Rock'ının Öncüsü; 70'ler-80'ler Atina
West Africa|Tel/Enstrümantal; Gelenek|Dinginlik|Odak; Gece|Kora; Griot Geleneği; Gambiya-Kopenhag
Western Europe|Elektronik|Dinginlik|Dans; Odak; Yolculuk; Spor|Downtempo; Trip-Hop; Sample Bazlı
Latin America|Gelenek|Karanlık|Dans; Yolculuk|Son Jarocho; Jarana-Charango; Balkan Etkisi
Latin America|Caz; Nefesli/Bando|Neşe|Dans; Odak; Yolculuk|Brazilian Jazz; Samba; Carnival Brass
North America|Gelenek; Pop/Şarkı-Yazarı|Neşe|Yolculuk|Old-Time; Bluegrass; Country Roots; Yukon
Western Europe; North Africa|Pop/Şarkı-Yazarı; Vokal/Şarkı|Epik|Dans; Spor|Kabyle Folk; Alternative Pop; Multilingual Song; Indie Folk
Western Europe|Psikedelik/Rock; Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Fransız Psikedelyası + Techno Prodüksiyonu
Southern Africa|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Gece|Afro-Pop; Lusophone Song; Contemporary World Music
Caribbean|Vokal/Şarkı; Gelenek|Neşe|Dans|Garifuna Paranda; Punta; Real World
Levant|Elektronik|Dinginlik|Odak; Yolculuk|Arap Elektronik; Tarab Füzyon; Vokal + Prodüksiyon
Western Europe|Pop/Şarkı-Yazarı; Caz|Melankoli|Gece|Akustik Pop; Vokal Armonisi; Swing; Halk Müziği; Latin Amerika ve Afrika Etkileri; Nefesliler
Mediterranean Europe|Blues|Melankoli|Gece|Gitar-Keman Düeti; Pizzica; Çöl Blues Etkisi
Iberia; North America; Latin America; Eastern Europe; Fusion|Caz|Melankoli|Gece; Yolculuk|Jazz; Free Jazz; Nu Jazz; Latin Alternative
North Africa; Oceania; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak; Yolculuk|Ud Virtüozluğu; Arap Makamı + Caz; Sydney
Levant|Elektronik|Karanlık|Dans; Odak|Arap Elektronik; Alternatif Pop; Trip-Hop
Western Europe|Caz; Elektronik|Dinginlik|Odak; Yolculuk|Akustik Trio; Elektronika Estetiği; Manchester; Blue Note
Western Europe|Pop/Şarkı-Yazarı; Nefesli/Bando|Neşe|Spor|Alternatif Şanson; Sokak/Punk Tavrı; Bordeaux
Western Europe|Pop/Şarkı-Yazarı|Neşe|Yolculuk|French Folk; Chanson; Poetic Songwriting
Western Europe|Caz; Nefesli/Bando|Neşe|Dans; Odak; Spor|French Jazz Orchestra; Big Band; Film Music; Experimental Jazz
Western Europe; Balkans; Fusion|Nefesli/Bando; Elektronik|Neşe|Dans; Odak; Spor|Balkan Nefesli + Kulüp Ritimleri; Londra
Western Europe; Fusion|Elektronik; Reggae/Dub|Neşe|Dans|Mediterranean Electronica; Dub; Jungle; Gnawa; Sephardic Folk
Balkans; Mediterranean Europe|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Laïko; Éntekhno
North America|Vokal/Şarkı; Blues|Neşe|Spor|Erken Caz/Ragtime Canlanması; Western Swing
Levant|Psikedelik/Rock|Neşe|Spor|Akdeniz Sörf Rock; Tuba + Gitar; Tel Aviv
Western Europe; Mediterranean Europe|Elektronik; Vokal/Şarkı|Neşe|Dans; Odak; Spor|Electro Swing; Downtempo; Vintage Pop
North America|Caz|Dinginlik|Odak; Gece|Space-Age Pop; Orkestra Şefi; Stüdyo Deneyleri
East Africa|Caz|Dinginlik|Gece; Yolculuk|Sudan Şarkısı; Caz; Afro-Füzyon
Mediterranean Europe|Caz; Nefesli/Bando|Neşe|Yolculuk|Napoli Geleneği + Caz; Politik Şarkı; Napoli
Caribbean|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Odak; Gece; Yolculuk|Çello-Vokal; Latin Oda Müziği; Paris
Anatolia|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Gece|Turkish Alternative Pop; Singer-Songwriter
East & Central Asia|Gelenek; Tel/Enstrümantal|Karanlık|Odak; Yolculuk|Yanggeum + Geomungo + Davul; Çağdaş Kore Topluluğu
Southern Africa|Elektronik|Karanlık|Dans; Odak; Yolculuk|Amapiano
Balkans|Nefesli/Bando|Neşe|Dans; Odak|Balkan Brass; Romani Wedding Band; Instrumental
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Gece|Contemporary British Folk; Acoustic Duo
Balkans|Vokal/Şarkı; Gelenek|Melankoli|Gece|Sevdah; Mostar Repertuvarı
Eastern Europe; Latin America; Fusion|Caz; Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Nu Jazz; Afro-Brazilian Electronic
Iranian World; Oceania; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Fars Cazı; Tar + Caz Kadrosu; Sydney Sahnesi
North America|Gelenek; Caz|Melankoli|Yolculuk|Celtic Fusion; Canadian Roots; Jazz
North America; Iranian World; Fusion|Klasik/Modern Klasik; Elektronik|Dinginlik|Dans; Odak; Yolculuk|Santur + Keman + Bas + Perküsyon; Doğaçlama Kuartet
Oceania; South Asia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans|Pencap Vokali + Çağdaş Pop; Melbourne
Latin America|Elektronik|Neşe|Dans; Spor|Dijital Cumbia; ZZK Records; Buenos Aires
West Africa; Western Europe; Fusion|Tel/Enstrümantal|Dinginlik|Dans; Yolculuk|Kora-Keman; Senegal Vokali; Avrupa Halk Müziği
Balkans|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak; Spor|Makedon/Balkan Repertuvar; Çoklu Enstrüman
North America; Iranian World; South Asia; Balkans; Fusion|Klasik/Modern Klasik; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Kamancheh; Tabla; Balkan Ensemble; Chamber Fusion
Southeast Asia|Psikedelik/Rock; Gelenek|Neşe|Dans; Spor|Khmer Rock; Psychedelic Pop; Garage Rock
North America|Funk/Soul; Caz|Melankoli|Yolculuk|Soul Jazz; Psychedelic Soul; Vocal Funk; Jazz Rock
Levant|Elektronik|Neşe|Dans; Odak; Spor|Funk; Video-Mashup Prodüksiyon; Elektronik Soul
Western Europe; Caucasus; Fusion|Pop/Şarkı-Yazarı|Neşe|Dans|Fransız Indie Pop; Ermeni Folk Konuğu
Iranian World|Gelenek; Elektronik|Melankoli|Gece|Persian Modal; Ambient; Tribal Percussion
North America|Psikedelik/Rock; Blues|Neşe|Dans; Yolculuk|Garage/Blues Rock; The White Stripes; Nashville
North America|Gelenek|Neşe|Dans; Yolculuk|Indie Folk; Keman; Islık ve Loop
North America; Balkans; Fusion|Nefesli/Bando; Hip-Hop/Rap|Neşe|Dans; Gece|Balkan Brass; Klezmer; Slavic Punk; Latin Hip-Hop
Western Europe|Gelenek; Elektronik|Dinginlik|Odak; Yolculuk|Breton Dance; Clarinet/Bass; Electronic Folk
Eastern Europe|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Odak|Hussar Funk; Cinematic Funk; Eastern Groove
Northern Europe|Psikedelik/Rock; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|İsveççe Psych-Pop; Stockholm
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Dans|Dutch-Language Kleinkunst; Folk-Pop; Avant-Garde
Western Europe; Anatolia; Fusion|Psikedelik/Rock; Elektronik|Neşe|Dans; Spor|Anadolu-Fransız Psych-Pop; Mr Bongo; Londra
Anatolia|Elektronik|Neşe|Dans; Odak; Spor|Türkçe Elektronik/Dans; İstanbul
Western Europe|Elektronik|Dinginlik|Odak|Trip-Hop/Downtempo; Marsilya
Western Europe|Elektronik|Neşe|Dans; Odak; Yolculuk|Trip-Hop; Lounge; Downtempo
Anatolia|Psikedelik/Rock; Elektronik|Romantik|Dans; Odak; Spor|Anadolu Psikedelyası; Elektro Saz
North America|Caz|Romantik|Gece; Yolculuk|Retro Pop Orkestrası; Çok Dilli Kabare; Latin Caz
Iranian World; Western Europe; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Gece|Fars Geleneksel/Halk Müziği; Setar; Vokal; Çağdaş Yorum; Viyana Merkezli Diaspora
Eastern Europe|Elektronik; Gelenek|Neşe|Dans|Ethno-Futurism; Tribal Electronica; Psychedelic Folk; Russian Lyrics
Western Europe; Northern Europe; Fusion|Tel/Enstrümantal|Epik|Dans; Odak; Gece; Yolculuk|Gitar; Nyckelharpa; Çok Kültürlü Yaylı Topluluğu
Western Europe|Elektronik; Klasik/Modern Klasik|Melankoli|Odak; Gece; Yolculuk|Ethereal/Neoklasik; Erken Müzik; Ortadoğu-Balkan Ödünçlemeleri; 4AD
Iberia|Psikedelik/Rock|Karanlık|Yolculuk; Spor|Flamenco Rock; Stoner; Psikedelik
East Africa|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk; Spor|Beja Geleneği; Kızıldeniz Blues; Tambo-Gitar
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Çağdaş Chanson; İspanyolca-Fransızca
Iranian World|Pop/Şarkı-Yazarı|Melankoli|Gece|Persian Pop; Modern Folk
Western Europe|Elektronik; Hip-Hop/Rap|Neşe|Yolculuk|Trip-Hop; Electro Swing; Nu Jazz; Sample-Based Beats
Balkans|Nefesli/Bando|Neşe|Dans; Odak|Balkan Brass; Roman Düğün Müziği; Çoček
Anatolia|Elektronik|Dinginlik|Dans; Odak; Gece; Yolculuk|Elektronik; Anadolu Psychedelia; Organik House
West Africa; Western Europe; Fusion|Blues; Psikedelik/Rock|Neşe|Dans; Odak|Bamako Çöl Bluesu; Songhai Gitarı; Rock Prodüksiyonu
West Africa; North Africa|Gelenek; Blues|Melankoli|Gece; Yolculuk|Guembri/Gnawa + Senegal Hattı; Dakar-Paris
Western Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Enstrümantal; Dans ve Tiyatro Müziği; Minimal Folk
West Africa|Blues|Melankoli|Dans; Yolculuk; Spor|Tuareg Assouf; Çöl Blues; Gitar Geleneği
Latin America|Gelenek; Psikedelik/Rock|Melankoli|Dans; Yolculuk; Spor|Peruvian Cumbia; Chicha; Psychedelic Tropical Rock
North America|Tel/Enstrümantal|Karanlık|Odak; Gece; Yolculuk|Appalachia Fiddle; Lap Steel; Drone-Folk
Western Europe|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Ronroco; Ambient/Neoklasik; Markus Sieber Projesi
Northern Europe|Gelenek; Elektronik|Melankoli|Gece|Talharpa (hiiu kannel) + Efektler; Tallinn
East & Central Asia|Psikedelik/Rock|Epik|Dans; Spor|Çin Azınlık Halk Müziği; Folk-Rock; Yunnan Geleneği
Anatolia; Latin America; Fusion|Elektronik; Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|'Tropical Anatolia'; Anadolu Ezgileri + Tropikal Ritim
North Africa|Tel/Enstrümantal; Psikedelik/Rock|Neşe|Dans; Odak; Spor|Elektrikli Gitar + Arap Makamı; 70'ler Kahire/Beyrut
Iberia|Caz; Elektronik|Karanlık|Dans; Odak; Yolculuk|Sefarad Cazı; Akdeniz Ezgileri; Serbest Doğaçlama; Ambient/Dub
Western Europe; Latin America; Fusion|Pop/Şarkı-Yazarı; Klasik/Modern Klasik|Neşe|Dans; Spor|Brazilian Pop; Welsh Orchestra; Tropicalia
North Africa|Vokal/Şarkı|Epik|Gece; Yolculuk|Sahrawi Müziği; Haul; Politik Şarkı
Oceania; Iranian World|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk; Spor|Caz Kontrbası + Fars Modal Müziği; Perth
Balkans|Psikedelik/Rock|Neşe|Spor|Yugo-Rock; Balkan Kabare; Politik Satır
Balkans|Nefesli/Bando|Neşe|Dans|Balkan Brass; Roman Nefesli Orkestrası
Western Europe|Nefesli/Bando; Gelenek|Neşe|Dans; Yolculuk|Klezmer; Amsterdam Sahnesi
Northern Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Dans; Odak; Gece|Oslo Yeraltı Sahnesi; Çok Disiplinli Şarkı Yazarlığı
Caucasus; Western Europe; Anatolia; Fusion|Hip-Hop/Rap|Romantik|Gece; Yolculuk|Ermeni Halk Müziği; Anadolu Füzyon; Hip-Hop
Western Europe|Funk/Soul; Elektronik|Neşe|Dans|Neo-Soul; Afro-Latin Groove; Balkan Klarnet; Hip-Hop/Drum'n'Bass; Bordeaux
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|French Chanson; Alternative Rock; Acoustic Ensemble
Western Europe|Klasik/Modern Klasik|Dinginlik|Odak|Kontrabas; Erken Müzik; Doğu-Batı Kamara Füzyonu
West Africa|Vokal/Şarkı; Funk/Soul|Melankoli|Yolculuk|Wolof Vokal; Perküsyon; Montreal
North Africa|Psikedelik/Rock|Melankoli|Odak|Tuareg Rock; Çöl Blues; Assouf
West Africa|Tel/Enstrümantal; Blues|Neşe|Yolculuk|Ngoni; Griot Geleneği; Afro-Blues
Latin America|Funk/Soul; Caz|Neşe|Odak|Afrobeat; Enstrümantal; São Paulo
North America; Latin America|Psikedelik/Rock; Elektronik|Karanlık|Odak; Yolculuk|Tropikal Fütürizm; Kolombiya Kökeni; New York
Latin America|Gelenek; Vokal/Şarkı|Epik|Dans; Yolculuk|São Paulo Deneysel Sahnesi; Afro-Brezilya Öğeleri
Northern Europe|Vokal/Şarkı; Gelenek|Dinginlik|Gece|Norveç Halk Şarkısı; Kulokk/Çağırma Geleneği
Iberia|Vokal/Şarkı|Neşe|Dans; Yolculuk|Yeni Flamenko; Roman Vokal; Nuevo Flamenco Kurucuları
West Africa; Western Europe; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Kora + Kelt Arpı İkilisi
Anatolia|Tel/Enstrümantal; Gelenek|Melankoli|Dans; Odak; Yolculuk|Bağlama/Saz; Anadolu Makam Geleneği
Levant|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Klasik Arap Vokali; Tarab; Caz-Flüt
Western Europe; West Africa; Fusion|Elektronik; Gelenek|Melankoli|Dans|African Electronic; Cabo Verdean Song
Western Europe|Gelenek; Nefesli/Bando|Neşe|Yolculuk|Klezmer + Balkan; Sokak Bandosu; Amsterdam
Mediterranean Europe; Western Europe; Fusion|Caz; Pop/Şarkı-Yazarı|Dinginlik|Dans; Odak|Caz Davulları; Akdeniz Şarkıcı-Şarkı Yazarlığı; Folk-Blues; Akustik İşbirliği
Levant|Elektronik|Neşe|Yolculuk|Arap Elektronik; Alternatif Pop; Trip-Hop
Western Europe; Anatolia; Caucasus; West Africa; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Gece|Karanlıkta İcra; ~30 Enstrüman (hoddu–saz); Muziekpublique; Brüksel
Latin America|Psikedelik/Rock|Neşe|Odak|Tropikal Psikedeli; Surf-Funk; Bogotá
North America|Pop/Şarkı-Yazarı|Melankoli|Gece|Amerikan Folk; Sade Akustik; Portland
North America|Blues; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Blues + Klezmer + Tuba; New York
Balkans|Vokal/Şarkı|Epik|Yolculuk; Spor|Yunan Alternatifi; 'Largo' EP (2016); Ahos Etiketi
North America|Caz|Dinginlik|Odak; Gece; Yolculuk|Caz Füzyon Gitarı; Flamenko Etkisi
Levant|Psikedelik/Rock|Romantik|Gece|Folk-Blues; Rock; İki Cinsiyetli Vokal Tonu
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans|Fransız Şansonu; Şarkıcı-Şarkı Yazarı; Halk Müziği; Blues
Eastern Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Gece|İnternet Pop; Satırik Şarkı; Rus Gitar Geleneği
Western Europe|Psikedelik/Rock; Elektronik|Neşe|Dans; Odak; Yolculuk|Kozmik Psych/Krautrock; Viyana
Western Europe|Klasik/Modern Klasik|Melankoli|Odak; Gece|Minimalist Piyano; Film Müziği; Breton Folk
Eastern Europe|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Balkan/Roman ve Ortadoğu Repertuvarı; Szczecin
Western Europe; Eastern Europe; Balkans; Fusion|Gelenek; Vokal/Şarkı|Neşe|Spor|Roma Music; Manele; Film Music
Western Europe; Arabian Peninsula|Caz; Elektronik|Dinginlik|Odak; Yolculuk|Arabic Jazz; Trumpet; Flugelhorn; Electronic Sound Design; Quarter-Tone Modes
North America; Iranian World; Fusion|Elektronik|Karanlık|Odak|Psikedelik Beat; Sample Prodüksiyon; Fars Etkisi
Western Europe|Blues; Psikedelik/Rock|Neşe|Dans|Garage Blues; Lo-Fi; Thee Headcoatees Çevresi
Caucasus|Gelenek; Vokal/Şarkı|Neşe|Dans|Çerkes/Adige Müziği; Kafe Dansı; Laperise; Kabardey/Adige Repertuvarı
Levant|Elektronik; Reggae/Dub|Dinginlik|Dans|Dünya Müziği / Dub Prodüksiyonu
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Dans|Sea Shanty; Maritime Folk; A Cappella
Mediterranean Europe|Gelenek; Elektronik|Neşe|Yolculuk|Salento/Pizzica Hattı; Çağdaş Düzenleme
Anatolia; Iranian World; Balkans; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Bağlama; Lyra; Persian Percussion; Modal Ensemble
East & Central Asia; Oceania; Fusion|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|Pacific Island Music; Taiwanese Indigenous
Northern Europe|Gelenek|Karanlık|Odak|Fin Halk Dans Müziği; Polkka; Kadın Kadro
Western Europe|Caz|Neşe|Dans; Odak|Manouş Caz Üçlüsü; Doğu Ezgileri
Levant; Arabian Peninsula|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Yemen-Yahudi Kadın Repertuvarı + Elektronik
Latin America|Gelenek|Melankoli|Gece|Akustik Folk; Latin Songwriter
North Africa|Vokal/Şarkı|Neşe|Dans|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
Western Europe|Psikedelik/Rock|Neşe|Dans|French Art Rock; Cabaret Noir; Chamber Rock
Balkans|Psikedelik/Rock; Nefesli/Bando|Neşe|Dans; Odak|Balkan Rock; Film Müziği; Kusturica Çevresi
Balkans|Caz|Melankoli|Yolculuk|Sevdalinka; Bosna Halk Müziği; Caz Düzenleme
Western Europe|Psikedelik/Rock|Karanlık|Yolculuk; Spor|Arabian Fuzz; Psikedelik Rock; Maghreb Füzyon
Latin America|Gelenek; Psikedelik/Rock|Neşe|Dans; Spor|Marimba; Cumbia; Punk; Ska
South Asia; Western Europe; Fusion|Tel/Enstrümantal|Epik|Dans|Racastan Roman Müziği; Kalbeliya Dansı; Gitar-Oud Füzyon
Caucasus|Vokal/Şarkı|Romantik|Gece|Gürcü Halk Müziği; Panduri; Vokal Üçlüsü
Western Europe|Elektronik; Caz|Melankoli|Odak|Elektronik-Caz İkilisi; Atom™ + Burnt Friedman
West Africa; Iberia; Fusion|Tel/Enstrümantal|Dinginlik|Dans; Gece; Yolculuk|Kora-Çello Düeti; Griot Geleneği; İber Vokal
Levant|Klasik/Modern Klasik|Dinginlik|Dans; Odak|Oud Üçlüsü; Arap Klasik Müziği
Iberia|Vokal/Şarkı|Melankoli|Dans; Gece|Portekiz Oda-Folk; Teresa Salgueiro; Lizbon
North Africa|Klasik/Modern Klasik; Vokal/Şarkı|Melankoli|Gece|Opera Sopranosu + Arap-Endülüs Repertuvarı
Western Europe|Tel/Enstrümantal|Dinginlik|Gece|Kaval; Balkan-Anadolu Nefesli Geleneği
Mediterranean Europe; Fusion|Caz; Elektronik|Karanlık|Odak; Yolculuk; Spor|Doğaçlama Enstrümantal; Caza Yakın Füzyon; Elektronik/Deneysel; Londra-Lefkoşa
Western Europe; East & Central Asia; Fusion|Tel/Enstrümantal; Gelenek|Neşe|Yolculuk|Üç Şamisen; Tsugaru Geleneği; Berlin
Levant; Western Europe; Balkans; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Mediterranean Folk; Arabic Song; Romani Music; Swing; Oud; Buzuq
Western Europe|Elektronik; Gelenek|Melankoli|Dans; Spor|Gaelce Vokal + Elektronik; Skye Sahnesi
Southern Africa|Klasik/Modern Klasik|Romantik|Odak; Gece; Yolculuk|Klasik Gitar; Kora Transkripsiyonu; Akustik Solo
Iberia|Vokal/Şarkı; Gelenek|Melankoli|Dans; Yolculuk|İspanyol Folk Canlanması; Jota / Charro / Fandanguillo / Pasodoble; Pandereta
Latin America; Iberia|Funk/Soul; Gelenek|Neşe|Yolculuk|Kolombiya-İspanya Hattı; Afro-Latin Funk; 70'ler
West Africa; Western Europe; Fusion|Tel/Enstrümantal; Vokal/Şarkı|Dinginlik|Gece|Kora + Gitar/Vokal; Mande – Akdeniz Folk Buluşması
Levant; North America; Fusion|Caz; Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk|Caz; Akdeniz; Surf Rock; Afrobeat; Afro-Küba; Etiyopya Cazı; Orta Doğu Groove'ları
Western Europe; Latin America; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece|Enstrümantal Çöl Gitarı; İki Kardeş; Easy Eye Sound
Western Europe|Elektronik|Melankoli|Dans; Yolculuk|Krautrock Etkili Elektronik; Düsseldorf
North America; Western Europe; Fusion|Elektronik|Melankoli|Dans|Organic House; Tribal House; Downtempo
Caribbean|Reggae/Dub|Neşe|Gece; Yolculuk|Roots Reggae; The Gladiators Çevresi; Kingston
Latin America; Western Europe; Mediterranean Europe; Fusion|Nefesli/Bando; Gelenek|Neşe|Yolculuk|Mexican Brass; Klezmer; Balkan Music; Street Band
Western Europe; Fusion|Caz; Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Naylon Telli Gitar; Çello; Caz; Klasik; Akdeniz/Dünya Müziği Doğaçlaması
Western Europe|Caz|Epik|Odak; Yolculuk|Ermeni Ezgileri; Caz Piyanosu; Şarkı Yazımı
East Africa|Vokal/Şarkı|Neşe|Dans|Maloya; Trans Vokal; Ritüel Müzik
Western Europe|Reggae/Dub|Karanlık|Dans; Odak; Yolculuk|Deneysel Dub; On-U Sound; Perküsyon Katmanları
Iberia|Vokal/Şarkı; Gelenek|Neşe|Dans|'Tasca Beat'; Portekiz Meyhane Repertuvarı; Lizbon
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans; Odak|Sade İndie Folk; İngilizce Söz; Lyon
Iberia|Caz|Romantik|Odak; Gece; Yolculuk|Caz Vokal; Trombon; Bossa Nova-Katalan Folk
Western Europe|Caz; Elektronik|Melankoli|Yolculuk|Caz-Breaks Prodüksiyonu; Tru Thoughts; Brighton
Western Europe|Gelenek|Neşe|Dans; Odak; Spor|Diyatonik Akordeon; Bal Musette; Fransız Halk Müziği
Western Europe; Mediterranean Europe|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Electro Swing; Downtempo; Vintage Pop
Western Europe; West Africa; Fusion|Tel/Enstrümantal; Pop/Şarkı-Yazarı|Dinginlik|Gece|Kora + Gitar; Fransız Şarkı-Yazarı
Western Europe|Elektronik; Funk/Soul|Neşe|Dans; Odak; Spor|Trip-Hop / Hip-Hop + Elektro; Chinese Man Kurucusu; Cavaquinho; Güney Amerika-Hindistan-Türkiye Dokuları
West Africa|Caz; Gelenek|Dinginlik|Dans; Odak|Kora + Piyano + Perküsyon Üçlüsü
Western Europe; Levant; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Electro-Pop; World Dance; Global Groove
Western Europe|Funk/Soul|Melankoli|Dans|Anonim Kolektif; Post-Punk Etkili Soul; Londra
Eastern Europe; Iberia; Fusion|Vokal/Şarkı; Gelenek|Romantik|Gece|Sefarad/Ladino Repertuvarı; Çek Topluluğu
Anatolia|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans; Yolculuk|Organic House; Downtempo; Turkish Vocal; Anatolian Electronic
Western Europe; Balkans; Fusion|Elektronik; Reggae/Dub|Neşe|Dans; Odak|Balkan Dub/Elektronik; Viyana
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Fransız Alternatif Şanson; Tiyatral Proje
West Africa|Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Afro-Küba; Senegal Pop; Mbalax Öncesi
Latin America|Elektronik|Dinginlik|Dans; Odak; Yolculuk; Spor|Deneysel Latin Elektroniği; Organic House; Milonga
Latin America; Mediterranean Europe|Vokal/Şarkı; Elektronik|Neşe|Dans; Yolculuk|Bossa Nova + Nu-Jazz; Roma Yerleşik
North America; Western Europe|Funk/Soul|Melankoli|Gece|Gospel-Soul Üçlüsü; Los Angeles/Londra
Eastern Europe; Western Europe; Northern Europe|Caz; Gelenek|Melankoli|Gece|Ukrainian Folk; Vocal Jazz; Contemporary Quintet
North America|Gelenek; Vokal/Şarkı|Melankoli|Dans|Bluegrass; Americana; Flatpicking Guitar
Western Europe|Caz|Melankoli|Yolculuk|Folk-Hop; Swing; Sokak Cazı
Western Europe|Funk/Soul|Neşe|Dans|Raw Soul; Rhythm and Blues; Garage Funk
Anatolia; Western Europe; North America; Caribbean; Fusion|Gelenek; Blues|Melankoli|Dans; Yolculuk|Kurdish Soul; Anatolian Folk; Raw Blues; Improvisation
West Africa; Levant; Fusion|Blues; Tel/Enstrümantal|Dinginlik|Dans; Odak; Yolculuk; Spor|Desert Blues; Acoustic Piano; Improvisation; Piyano-Gitar Doğaçlaması; İsrail-Mali Füzyonu; İsrail-Mali Füzyon; Piyano-Gitar Düeti; Guitar; West African Groove
West Africa; Western Europe; Fusion|Gelenek; Nefesli/Bando|Melankoli|Dans|Kora; Mandingue; Marching Band
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|İngiliz Politik Şarkı-Yazarlığı; Hiciv
Levant; Western Europe; Fusion|Tel/Enstrümantal; Caz|Romantik|Odak; Yolculuk|Kanun + Piyano; Çağdaş Oda Müziği
Western Europe|Blues|Melankoli|Dans|Amerikana; Folk-Blues; Çok Uluslu Grup
Iberia|Klasik/Modern Klasik|Melankoli|Yolculuk|Polifonik Vokal Düeti; Akdeniz İlahileri; Erken Müzik; Catalonia
Southern Africa; Northern Europe; Fusion|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Cross-Cultural Pop Collaboration
Western Europe|Caz|Melankoli|Yolculuk|Dünya-Caz Kuarteti; Akdeniz/Balkan Ezgileri; Münih
North America|Caz|Neşe|Spor|Manouş/Swing Cazı; Denver
Western Europe; Mediterranean Europe|Gelenek; Caz|Neşe|Dans|Avrupa Folkü + Modal Caz; Flamenko, Arap Müziği, Ethio-Jazz; Flüt-Arp-Akordeon-Keman
Levant|Pop/Şarkı-Yazarı; Blues|Melankoli|Gece; Yolculuk|Folk-Rock; İngilizce Söz; Tel Aviv
Western Europe|Funk/Soul; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Noir Soul; Boğuk Vokal; Eindhoven
Iranian World|Elektronik|Romantik|Spor|Fars Pop; Elektronik; Diaspora Funk
North America; Western Europe|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Deneysel Folk; Manchester Yerleşik
Latin America|Gelenek; Klasik/Modern Klasik|Melankoli|Gece|South American Song; Baroque Instruments; Voice/Flute
Latin America|Gelenek|Melankoli|Dans; Gece|Müzik Llanera; Arpa/Cuatro; Joropo
North America; Latin America; Fusion|Elektronik; Gelenek|Neşe|Dans; Odak|Afro-Brazilian Percussion; Global Bass; Pará Guitar; Electronic Roots
North America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Dans; Gece|Old-Time/Folk İkilisi; Vokal Harmoni; Nova Scotia
Western Europe; North America|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|İskoç Folkü + Erken Amerikan Halk Müziği; Banjo, Harmonika; 'Sorley the Bard'; Scotland, UK
East & Central Asia|Gelenek; Klasik/Modern Klasik|Neşe|Dans; Spor|Shakuhachi/Koto + Modern Düzenleme
Anatolia|Caz; Vokal/Şarkı|Dinginlik|Dans; Odak|Caz-Soul Vokali; İstanbul
Oceania|Vokal/Şarkı; Blues|Neşe|Dans|Country-Soul; Rockabilly; Güçlü Vokal
North America|Psikedelik/Rock|Epik|Gece|Amerikana; Noir Folk; Rockabilly-Blues
Western Europe|Gelenek; Caz|Romantik|Dans; Odak|Klezmer Klarneti; 'Yeni Yahudi Müziği'; Paris
Western Europe|Gelenek; Nefesli/Bando|Neşe|Gece|Balkan/Klezmer/Fransız Halk Ezgileri; Berlin Kolektifi
North Africa|Funk/Soul|Romantik|Dans; Yolculuk; Spor|Libya Disco-Funk; Soul; Reggae Etkisi
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Fransız Şansonı; Konuşur Gibi Vokal; Paris
West Africa; Southern Africa; Fusion|Tel/Enstrümantal|Epik|Spor|Fula Şarkısı; Akustik Gitar; Mbalax Kökleri
Levant|Caz|Neşe|Dans; Odak|Saksofon; Arap Ezgileri + Caz; Paris
Western Europe|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Dans; Gece|Çağdaş Şanson; Sintezleyici Düzenleme; Paris
North America; Latin America; West Africa; Fusion|Funk/Soul; Psikedelik/Rock|Neşe|Odak|Enstrümantal Gitar; Chicha/Tuareg/Son Montuno; New Orleans
North America|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Québécois Pop; Pop-Rock; Synth-Pop
North America|Gelenek; Funk/Soul|Neşe|Dans; Yolculuk; Spor|Sikh Kirtan; Punjabi; Bhangra; Brass
Iranian World|Elektronik; Psikedelik/Rock|Melankoli|Gece; Yolculuk|Experimental Electronic; Iranian Alternative
Western Europe|Caz; Gelenek|Melankoli|Dans; Odak; Yolculuk|Manouş Cazı + Dünya Ezgileri; Paris
Western Europe; South Asia; Fusion|Caz; Nefesli/Bando|Neşe|Odak; Spor|Geniş Kadro; Çok Kültürlü Londra; Hint/Çin/Bangladesh Gelenekleri
Western Europe; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans|Electro-Klezmer; Alternative Hip-Hop; Yiddish Song
West Africa|Blues; Tel/Enstrümantal|Neşe|Dans; Yolculuk|Çöl Blues; Songhai/Peul Ezgileri; Diré
Eastern Europe|Gelenek; Psikedelik/Rock|Neşe|Odak; Yolculuk|Cimbalom + Rock Kadrosu; Macar Halk Ezgileri
Eastern Europe|Gelenek|Neşe|Dans|Macar Roman Halk Müziği; Köy Müziği; Ağız Bası; Ev Eşyası Perküsyonu; Dans
Eastern Europe|Vokal/Şarkı; Gelenek|Neşe|Spor|Macar Roman (Lovari) Vokal Geleneği
Western Europe; Eastern Europe; Fusion|Nefesli/Bando; Gelenek|Neşe|Gece|Klezmer + Macar Tambura Geleneği
East Africa; North America|Caz; Funk/Soul|Melankoli|Dans; Odak; Yolculuk|Ethio-Jazz; Klavye/Besteci; Los Angeles
West Africa|Funk/Soul|Romantik|Yolculuk|Benin Afro-Funk; 70'ler; Yeniden Basım
Western Europe; Latin America; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Dans; Gece; Yolculuk|Samba de Raiz; Afro-Brezilya Perküsyonu; Fransız Şansonu; Karayip Beguine'i; New Orleans Nefesli Geleneği; Caz Doğaçlaması
Balkans; Eastern Europe; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Dans; Odak; Gece|Carpathian Folk; South Slavic; Tamburitza/Double Bass
Western Europe|Caz; Pop/Şarkı-Yazarı|Melankoli|Dans|Caz-Şanson; Retro Kadro; Paris
Western Europe|Gelenek; Psikedelik/Rock|Neşe|Dans|İngiliz Folk-Rock; Maddy Prior; 1969
Western Europe|Caz|Neşe|Dans; Odak; Yolculuk|Londra Caz Sahnesi; Afrobeat/Grime Etkileri
South Asia|Vokal/Şarkı; Elektronik|Melankoli|Yolculuk|Chitrali Folk; Khowar; Drum and Bass
Iranian World|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Kamançe Soloları; Radif; Kaçık Ölçüler; Kamançe–Perküsyon Düosu
Iranian World; Western Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece; Yolculuk|Setar/Daf + Zarb; Fars Modal Geleneği
West Africa; Western Europe; Fusion|Psikedelik/Rock; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Burkina Faso Griot Vokali + Belgçika Noise-Rock
North America; Latin America; Fusion|Funk/Soul; Caz|Neşe|Dans; Spor|Political Funk; Cumbia; Jazz; Ska; Cuban-American Vocal
Western Europe; East & Central Asia; Iranian World; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Odak; Gece|Moğol Gırtlak Vokali + Fars Perküsyon + Şansın
Western Europe; South Asia; Fusion|Psikedelik/Rock|Karanlık|Dans; Yolculuk|Hint-Viktorya Füzyonu; Sitar-Rock; Steampunk Estétik
Northern Europe; South Asia; North Africa; Iranian World; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Yolculuk|Hindustani Raga; Swedish Folk; Arabic Maqam; Sitar; Oud
West Africa|Gelenek; Vokal/Şarkı|Neşe|Gece; Yolculuk|Kora; Griot Geleneği; Dakar
West Africa; Iberia|Gelenek; Vokal/Şarkı|Neşe|Dans; Odak|Balafon; Griot Geleneği; Lizbon Yerleşik
Balkans|Elektronik|Melankoli|Gece|Yunan Retro Vokal; Downtempo; Sample Prodüksiyon
North America|Gelenek|Dinginlik|Odak; Gece|Kelt Kemanı; Step Dans; Vancouver
Latin America|Vokal/Şarkı; Gelenek|Dinginlik|Dans|Brezilya Müziği (MPB); Yerel Kadro
Mediterranean Europe; Fusion|Gelenek; Psikedelik/Rock|Neşe|Dans; Spor|Kıbrıs Halk Müziği; Doğu Akdeniz; Tzouras; Deneysel/Psikedelik; Dub; Ses Sanatı
Northern Europe|Gelenek|Epik|Odak|'Yükseltilmiş Tarih'; Ritual Folk; Eski Nordik Metinler
Western Europe; West Africa|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Oksitan + Batı Afrika Hattı; Çok Dilli Repertuvar
Western Europe|Gelenek; Nefesli/Bando|Neşe|Odak|İskoç Gaydası; Pipe Band; Modern Jig/Reel; Scotland
Northern Europe|Gelenek|Neşe|Spor|Çok Dilli Halk Şarkıları; Kantele/Jouhikko; Luciano Berio Folk Songs
Fusion|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Spor|Global Collaboration; Roots Music
Levant|Caz|Dinginlik|Odak; Yolculuk|Contrabas Caz; Akdeniz Cazı; World Jazz
Latin America|Gelenek; Klasik/Modern Klasik|Melankoli|Odak; Gece|South American Song; Baroque Instruments; Voice/Flute
Levant|Psikedelik/Rock; Gelenek|Neşe|Odak|Orta Doğu Psikedelyası; Türk Halk Müziği Uyarlaması; Surf-Funk; Ud; Irak-Arap Vokal Bağlamı
North America|Blues|Melankoli|Dans|Amerikana; Blues; Appalachia Vokal
Western Europe; Balkans; Fusion|Caz; Gelenek|Dinginlik|Dans; Odak; Gece|Jazz Manouche + Balkan/Klezmer
West Africa|Funk/Soul; Vokal/Şarkı|Neşe|Dans; Odak; Yolculuk; Spor|Malian Afro-Funk; Organ/Guitar; Vintage Groove
Iberia; Western Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Odak|Murcian Folk; Mandolin Family; Cello
Western Europe; Levant|Caz; Nefesli/Bando|Dinginlik|Odak; Gece; Yolculuk|Flüt/Ney; Arap Makamı + Caz; Paris
Western Europe|Gelenek; Nefesli/Bando|Melankoli|Gece; Yolculuk|Balkan/Klezmer/Fransız Halk Ezgileri; Berlin Kolektifi
East Africa|Vokal/Şarkı; Gelenek|Neşe|Dans; Yolculuk|Afar Song; A Cappella; Minimal Guitar; Oral Tradition
Western Europe; Anatolia; Fusion|Elektronik|Karanlık|Dans|Enstrümantal Elektronik; Asya-Ortadoğu Sample; Psikedelik
North Africa|Vokal/Şarkı|Melankoli|Yolculuk|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
North Africa|Psikedelik/Rock; Gelenek|Dinginlik|Dans; Odak; Gece; Yolculuk|Amazigh Rock; Ouarzazate; Tamazight Söz
Levant|Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Yolculuk|Hebrew Song; Folk; Live Looping; Falafel Pop
Balkans|Gelenek|Neşe|Dans|Bulgar Halk Müziği; Didgeridoo-Kaval; Dubstep-Drum'n'Bass
West Africa; Fusion|Gelenek; Caz|Neşe|Dans; Yolculuk; Spor|Kora; World Jazz; West African Song; Mande Geleneği; Griot Müziği
Latin America; Western Europe; East & Central Asia; Fusion|Caz; Gelenek|Melankoli|Gece; Yolculuk|Chamber Jazz; String Ensemble; Global Fusion
Western Europe|Gelenek; Tel/Enstrümantal|Neşe|Yolculuk|Scottish Fiddle; Shetland Tradition; Québécois Folk; Foot Percussion
Latin America|Gelenek; Elektronik|Neşe|Yolculuk; Spor|Millennial Cumbia; Tropicanibal; Psychedelic Tropical
North America|Vokal/Şarkı; Gelenek|Neşe|Dans|Appalachian Kadın Vokal İkilisi; Bluegrass Öncüleri
Caribbean; Latin America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans; Gece|Latin Alternatif; Kadın Vokal Düeti
Iranian World; West Africa; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Kamanche-Kora Düeti; Doğaçlama; Kamara Füzyon
Western Europe|Caz; Elektronik|Dinginlik|Dans; Odak|Akustik Trio; Elektronika Estetiği; Manchester; Blue Note
Northern Europe|Gelenek; Elektronik|Karanlık|Yolculuk|Fin Halk Müziği; Kantele + Elektronik; Helsinki
Levant|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Ney; Türk/Arap/Fars Ney Gelenekleri; Sufi Müziği
Anatolia; Western Europe; Fusion|Elektronik; Gelenek|Neşe|Dans; Yolculuk; Spor|Anatolian Folk; Electronic Beat; Studio Fusion
North Africa; Western Europe|Klasik/Modern Klasik; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Mısır Şarkı Geleneği + Yaylı Düzenleme; Scotland
Western Europe; Anatolia; Fusion|Elektronik; Psikedelik/Rock|Neşe|Dans; Spor|Kürt/Türk Müziği; Anadolu Psikedelyası; Türk Diskosu; Elektrikli Saz; Synth Groove
Northern Europe; Eastern Europe; North Africa; Fusion|Caz; Gelenek|Dinginlik|Gece; Yolculuk|Microtonal Folk-Jazz; Arabic Maqam; Norwegian Folk
Levant|Funk/Soul; Psikedelik/Rock|Neşe|Dans|Psych-Pop; Soul; Middle Eastern Psychedelia
North Africa|Pop/Şarkı-Yazarı|Neşe|Yolculuk; Spor|Nubya Müziği; Elektrikli Nubya Pop
West Africa|Gelenek; Blues|Neşe|Dans; Odak; Yolculuk|Malian Guitar; Manding; Bambara; Songhai; Afro-Blues
Oceania|Caz; Funk/Soul|Epik|Dans; Odak; Yolculuk; Spor|Afrobeat; Spiritual Jazz; Geniş Kadro; Adelaide
Latin America; Western Europe; Fusion|Psikedelik/Rock|Karanlık|Dans; Odak|Tropikal Psikedeli; Analog Prodüksiyon; Amsterdam
Levant|Reggae/Dub|Romantik|Dans; Spor|Reggae; Çöl Blues; Arap Rock
North Africa; Oceania; Fusion|Tel/Enstrümantal; Gelenek|Melankoli|Odak; Gece; Yolculuk|Oud; Yidaki; Australian Chamber Folk
Northern Europe|Caz; Tel/Enstrümantal|Melankoli|Yolculuk|Contemporary Jazz; Violin-led Ensemble
West Africa; Northern Europe; Fusion|Gelenek; Elektronik|Neşe|Dans; Spor|Senegalese Folk; Spiritual Jazz; Dub; Sufi Song; Experimental Production
Eastern Europe|Klasik/Modern Klasik; Caz|Melankoli|Gece; Yolculuk|Akordeon; Çağdaş Doğaçlama; Belarus
Iberia|Klasik/Modern Klasik|Melankoli|Dans; Spor|Polifonik Vokal Düeti; Akdeniz İlahileri; Erken Müzik; Catalonia
West Africa; Western Europe; Fusion|Gelenek; Caz|Melankoli|Gece; Yolculuk|Kora; World Jazz; Chamber Music
Latin America; Western Europe; North America|Tel/Enstrümantal; Psikedelik/Rock|Dinginlik|Odak|Instrumental Guitar; Latin Desert; Cinematic Surf
Northern Europe|Gelenek; Vokal/Şarkı|Neşe|Dans|North Norwegian Folk; Fiddle; Coastal Song
West Africa|Blues|Melankoli|Odak; Yolculuk|Mali Blues; Kayes Gitar; Solo Vokal-Gitar
North America|Gelenek; Vokal/Şarkı|Neşe|Dans|Bluegrass; Americana; Flatpicking Guitar
Mediterranean Europe|Vokal/Şarkı|Melankoli|Dans; Gece|Güney İtalya Vokali; Akdeniz Repertuvarı; Puglia
Western Europe; Southeast Asia; Fusion|Elektronik|Melankoli|Yolculuk|Deneysel Elektronik; İskoç-Burma Kökeni; Görsel İşler; Scotland
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Odak|Contemporary British Folk; Banjo; Bouzouki
Oceania|Vokal/Şarkı; Gelenek|Karanlık|Odak|Aborjin Vokali; Numbulwar; Wubuy Dili
North Africa; Iranian World; North America; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Rumi Poetry; Persian Setar; Arabic Vocal
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Dans|Çok Kültürlü Beşli; Manchester; Yaylı + Vokal
Balkans|Elektronik; Gelenek|Karanlık|Dans; Odak|Balkan Sample'ları + Elektronik; Canlı Kadro; Sofya
Iberia|Vokal/Şarkı|Melankoli|Dans|Çağdaş Fado; Lizbon
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Celtic Music; Irish Folk
Southeast Asia|Gelenek|Dinginlik|Gece|Sunda Repertuvarı; Eski Yazımla Kayıt
Balkans|Nefesli/Bando|Epik|Dans|Balkan Nefesli; Film Müziği; Wedding & Funeral Band
Western Europe|Caz|Neşe|Dans; Odak; Yolculuk; Spor|Jazz Manouche; Django Geleneği; New York
West Africa|Funk/Soul|Neşe|Odak; Yolculuk|Afrobeat; Highlife; Togolese Funk; Horn-Driven Groove
Latin America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Odak; Gece|Brazilian Pop; Samba; Contemporary Song
Western Europe|Pop/Şarkı-Yazarı; Reggae/Dub|Neşe|Dans; Spor|Çağdaş Chanson; Reggae/Ska; Toplumsal Söz
Iranian World|Gelenek; Tel/Enstrümantal|Romantik|Odak; Yolculuk; Spor|Kamançe Soloları; Radif; Kaçık Ölçüler; Kamançe–Perküsyon Düosu
Iberia|Nefesli/Bando|Neşe|Odak; Spor|Balkan Brass; Roman Müziği; Klezmer Füzyon
Levant|Elektronik; Caz|Karanlık|Yolculuk; Spor|Arap Şaabi; Free Jazz; Elektroakustik
Northern Europe; Iberia; Fusion|Gelenek; Caz|Melankoli|Gece|Nordic Folk; Iberian Folk; Improvised Ensemble
Iberia|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Katalan Vokali; Flamenko-Caz; Fado Etkisi
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Spor|Fransız Şarkıcı-Şarkı Yazarı; Şanson; Halk Müziği; Blues; Brezilya/Güney Amerika Etkileri
Oceania; Latin America; Fusion|Tel/Enstrümantal; Gelenek|Dinginlik|Odak; Gece|Brezilya Gitarı; Choro; Samba; Cavaquinho; Viola Caipira; Latin Caz
Latin America; Western Europe; Fusion|Funk/Soul; Elektronik|Neşe|Dans; Odak; Spor|Venezuela Psikedelik Funkı; Tamborera/Gaita; Elektro-Karayip
North America|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Gece|Americana Vokali + Film Müziği Yaylıları
Western Europe|Caz|Dinginlik|Odak; Gece|Jazz Manouche; Django Geleneği; New York
Western Europe|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Broadside Ballads; Experimental Folk; Vocal Trio
Balkans|Caz|Romantik|Spor|Sevdah; Roman Müziği; Balkan Caz
Balkans|Gelenek|Melankoli|Gece|İşkodra Kent Müziği; Arnavut Vokal Geleneği; Serenat
North America|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Yidiş/Klezmer; Montreal; Post-Rock Bağlantısı
Western Europe|Elektronik|Dinginlik|Gece; Odak|Enstrümantal Elektronik; Dünya Müziği Sample; Downtempo
Western Europe; West Africa; Fusion|Elektronik; Funk/Soul|Dinginlik|Yolculuk|Fransız Beat Prodüksiyonu + Senegal Vokali
Oceania|Vokal/Şarkı; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Yidiş/Klezmer; Weimar Kabare; Free Jazz; Big Band; Funk; Melbourne
North Africa; South Asia; Southeast Asia; North America; Latin America; Fusion|Caz; Gelenek|Neşe|Dans; Odak|World Jazz; Oud; Sitar; Fusion
Western Europe; East Africa; Fusion|Caz|Dinginlik|Dans; Odak; Yolculuk|Ethio-Jazz; Enstrümantal Funk
Northern Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|İskandinav Folk; Parmak Ucı Gitar; Sade Prodüksiyon
Mediterranean Europe; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Gece; Yolculuk|Keman + Ud + Ortaçağ Telleri; Modal Repertuvar
Western Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|İngiliz Halk Müziği; East Anglia Halk Müziği; Mandolin; Banjo; Old-Time; Bluegrass
Caribbean|Vokal/Şarkı|Epik|Dans; Spor|Haiti Folk; Alternatif Gitar; Kreol Şarkı
Iberia|Vokal/Şarkı|Melankoli|Gece|Ladino Şarkısı; Sefarad Repertuvar; Flamenko Etkisi
Northern Europe|Elektronik|Karanlık|Dans; Odak; Spor|Elektro-Tango; Bandoneon + Prodüksiyon
Mediterranean Europe|Gelenek|Neşe|Dans; Yolculuk|Sicilya Halk Müziği; Tarantella-Mazurka; Çerçeve Davul-Akordeon
North America; West Africa; Fusion|Psikedelik/Rock; Gelenek|Epik|Yolculuk|Mali Vokali + Psikedelik Rock; Oakland
Iberia; Western Europe; Fusion|Elektronik; Gelenek|Dinginlik|Dans; Odak; Yolculuk|Portekiz Halk Müziği; Caz; Elektronik; Brezilya/Yeşil Burun/İspanya Etkileri
Western Europe|Caz; Elektronik|Neşe|Dans; Yolculuk; Spor|Elektro-Caz; Sinematik/Film Müziği; Dünya Füzyonu; Paris–İstanbul
Western Europe|Gelenek|Melankoli|Dans; Odak|İngiliz Geleneksel Şarkı; Ham Prodüksiyon; Londra
Iberia|Vokal/Şarkı; Gelenek|Neşe|Spor|İber Halk Vokali; Galiçya-Katalonya-Kanarya Hattı
Balkans|Klasik/Modern Klasik; Tel/Enstrümantal|Melankoli|Gece|Klasik Gitar; Yunan Besteci; Atina
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Çağdaş Chanson; Mizahi/Anlatıcı Söz
Eastern Europe|Tel/Enstrümantal|Epik|Dans; Odak; Spor|Solo Akustik Gitar; Sokak/Gezgin İcra; Flamenko-Balkan Etkileri
Caribbean|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk; Spor|Guyana Folk-Funk; Kwe-Kwe/Afro-Karayip Perküsyon
Western Europe|Psikedelik/Rock; Elektronik|Karanlık|Odak; Yolculuk|Kozmik Psych/Krautrock; Viyana
East Africa|Vokal/Şarkı|Karanlık|Dans; Odak; Spor|Acholi Geleneği; Sürgün Şarkıları; Real World
Mediterranean Europe|Reggae/Dub|Melankoli|Dans; Yolculuk|Napoli Dub; Trip-Hop; Akdeniz Elektronik
Western Europe|Caz; Elektronik|Dinginlik|Odak; Yolculuk|Gal Halk Ezgileri; Dub; Kontrbas
Levant|Caz|Neşe|Odak|Enstrümantal; Beyrut Kütüphane Müziği; Arap Caz-Funk
West Africa|Blues|Neşe|Dans; Spor|Kamele Ngoni; Mande Müziği; Çöl Blues
Western Europe; North Africa; Fusion|Elektronik|Karanlık|Yolculuk; Spor|Elektronik; Arap Vokal; Trip-Hop
Levant|Tel/Enstrümantal|Dinginlik|Odak; Gece|Ortadoğu Enstrümantal; Ney-Oud; Meditatif Füzyon
North America|Blues|Karanlık|Odak; Spor|Junkyard Blues; Deneysel Şarkı Yazımı; Kabare
Caucasus|Gelenek|Romantik|Dans|Ermeni Halk Müziği; Balkan Füzyon; Pop
Iranian World|Vokal/Şarkı; Elektronik|Karanlık|Odak|Fars/Ortadoğu Vokali; Santur; Niyaz & Vas
Latin America; North America; Fusion|Elektronik; Vokal/Şarkı|Epik|Dans; Yolculuk|Brezilya Yerli Vokali + Çağdaş Prodüksiyon
Arabian Peninsula; Levant|Blues|Melankoli|Gece; Yolculuk|Yemen Folk; Funk; Blues-Ortadoğu Füzyonu
Western Europe|Blues|Neşe|Dans|Downtempo; Electro Swing; Chillout Prodüksiyon
Eastern Europe|Gelenek; Vokal/Şarkı|Neşe|Dans|Balkan/Roman ve Ortadoğu Repertuvarı; Szczecin
North America|Psikedelik/Rock|Karanlık|Gece|Gotik Blues; Güney Amerikana; Punk Etkisi
Western Europe|Caz; Elektronik|Melankoli|Dans|Contemporary Jazz; Vocal Ensemble; Remix
North America|Caz|Karanlık|Odak; Yolculuk|Avangard Trompet; Fly or Die; Politik Söz
Western Europe|Elektronik; Psikedelik/Rock|Melankoli|Yolculuk|Noir Film Müziği Estetiği; Magazine/Bad Seeds Geçmişi
Mediterranean Europe|Psikedelik/Rock|Karanlık|Yolculuk|İtalyan Alternatif Rock; Post-Punk
Western Europe|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk; Spor|Manouche/Gypsy Jazz Gitar İkilisi
North America|Psikedelik/Rock|Neşe|Odak; Spor|Surf Rock; Retro Pop Orkestrası; İşbirliği
Western Europe|Caz; Elektronik|Dinginlik|Odak|Hang/Handpan; Ambient Caz; Londra
West Africa|Gelenek; Blues|Neşe|Dans; Spor|Mande Repertuvarı; Bamako; Institut des Jeunes Aveugles
Iranian World|Vokal/Şarkı|Epik|Yolculuk; Spor|Tanbur; Klasik İran Müziği; Kadın Vokali
Iranian World|Vokal/Şarkı|Melankoli|Gece|Klasik İran Müziği; Tanbur-Kemençe; Radif Geleneği
Western Europe; North Africa; Balkans; Fusion|Caz|Romantik|Gece|Pontus Kemençesi; Arap Vokal; Kamara Caz
Iberia|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Fado; Yeni Kuşak; Lizbon
Southeast Asia|Psikedelik/Rock|Karanlık|Spor|Endonezya Psychedelia; Dangdut-Funk; Enstrümantal
North America; Latin America; West Africa; Fusion|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Odak|Enstrümantal Gitar; Chicha/Tuareg/Son Montuno; New Orleans
Latin America; Western Europe; Fusion|Vokal/Şarkı|Melankoli|Gece|Kamara Folk; Çok Dilli Vokal Düeti
North America|Vokal/Şarkı; Funk/Soul|Neşe|Gece|Americana; Country-Soul; Gospel; Roots Music
Anatolia; Mediterranean Europe|Caz; Tel/Enstrümantal|Dinginlik|Odak|Contemporary Jazz; Guitar; Percussion; Double Bass; Woodwinds
Oceania; Western Europe|Pop/Şarkı-Yazarı; Blues|Melankoli|Yolculuk|Americana/Roots Şarkı-Yazarı Düeti
Northern Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Free/Spiritual Jazz; Klarnet; Oslo Kolektifi
Iberia|Gelenek; Vokal/Şarkı|Neşe|Spor|Iberian Oral Tradition; Wedding Song; Voice/Percussion
Mediterranean Europe|Gelenek; Vokal/Şarkı|Melankoli|Dans|Güney İtalya Halk Repertuvarı; Kalabriya Hattı
Iberia|Psikedelik/Rock|Neşe|Dans; Spor|Flamenco Rock; Stoner; Psikedelik
South Asia|Tel/Enstrümantal|Neşe|Dans; Odak; Spor|Peştun Rubab; Enstrümantal Rock; Khyber Geleneği
Anatolia; Balkans|Gelenek|Melankoli|Yolculuk|Akordeon; Balkan-Anadolu Repertuvarı; Arşivcilik
Western Europe; Levant; Fusion|Caz; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Ud; Ney; Arap Makamı; Çağdaş Caz; Arap-Caz Füzyonu
Iranian World|Caz; Pop/Şarkı-Yazarı|Melankoli|Gece|Fars Alternatif; Caz-Rock; Akordeon-Piyano
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Spor|Anti-Folk/İndie; İngilizce Söz; Paris
Latin America; Balkans; Fusion|Gelenek; Caz|Neşe|Dans; Spor|Forró; Balkan Kopanitsa; Brazilian Jazz
West Africa|Tel/Enstrümantal; Vokal/Şarkı|Dinginlik|Odak; Gece; Yolculuk|Gine Gitarı; Mande Ezgileri; Paris
Southern Africa|Elektronik|Karanlık|Yolculuk; Spor|Deneysel Elektronik; Queer Sanat; FAKA Çevresi
East Africa|Funk/Soul; Gelenek|Neşe|Dans|Darfur Kökenli; Sudan Groove'u; Diaspora
West Africa|Gelenek|Neşe|Dans; Yolculuk; Spor|Kologo (İki Telli Lut); Frafra Geleneği; Bolgatanga
West Africa|Blues; Tel/Enstrümantal|Dinginlik|Yolculuk|Çöl Blues; Songhai Gitarı; Niafunké
North America|Gelenek|Neşe|Dans|Çağdaş Bluegrass; Boston
West Africa; North America; Fusion|Blues|Melankoli|Dans; Spor|Tuareg Assouf; Çöl Blues; Amerikana Konuğu
Western Europe|Gelenek|Dinginlik|Gece|Çok Dilli Folk; Chanson; Ada Müzikleri Füzyonu
Mediterranean Europe|Caz; Vokal/Şarkı|Neşe|Spor|Napoli Geleneği + Caz; Politik Şarkı
Balkans|Vokal/Şarkı|Melankoli|Gece|Rebetiko; Sefarad/Judeo-İspanyol; Balkan Repertuvarı
Northern Europe|Gelenek|Neşe|Yolculuk|Fin Arkaik Folkü; Jouhikko ve Akordeon
Western Europe|Psikedelik/Rock|Melankoli|Dans|Art-Rock; Gitar + Dramatik Vokal; Londra
Anatolia|Elektronik|Neşe|Yolculuk; Spor|Anadolu Psychedelia; Funk; Elektronik Folk
North America|Gelenek; Blues|Dinginlik|Odak|Siyah Telli Çalgı Geleneği; Banjo/Kemik; 'American Songster'
Anatolia|Elektronik|Dinginlik|Dans; Odak; Yolculuk|Elektronik; Anadolu Psychedelia; Organik House
Balkans|Gelenek|Romantik|Dans; Yolculuk; Spor|Roman Müziği; Manele; Läutari Geleneği
East & Central Asia|Gelenek|Epik|Dans; Odak; Yolculuk|Kazak Etno-Folk; Dombra/Kobyz; Gırtlak Vokali
Eastern Europe; Western Europe|Klasik/Modern Klasik; Caz|Dinginlik|Odak; Gece; Yolculuk|Neoclassical Piano; Ambient Jazz; Guest Saxophone
Western Europe; East Africa; Fusion|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk; Spor|Ethio-Funk; 1960'lar/70'ler Etiyopya Funk'ı; Soul; Addis Ababa Groove
West Africa|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Moritanya Aile Topluluğu; Sahra Repertuvarı; 70'ler
North Africa; Western Europe; Fusion|Gelenek|Neşe|Yolculuk; Spor|Fas Chaabi/Aita; Cinsiyet Rollerinin Tersine Çevrilmesi
Iberia|Gelenek; Tel/Enstrümantal|Epik|Yolculuk; Spor|Flamenko Gitarı + Cante; Kardeş İkilisi
Western Europe; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Modern Halk Müziği; Çok Sesli Kadın Vokalleri; Akustik Çalgılar; Çok Dilli Dünya Müziği Repertuvarı
North Africa; Western Europe; Fusion|Caz; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Arapça Spoken Word + Avrupa Cazı; Trompet-Saksofon
North Africa|Funk/Soul|Romantik|Dans; Yolculuk; Spor|Libya Disco-Funk; Soul; Reggae Etkisi
Northern Europe; East Africa; West Africa; Fusion|Caz; Funk/Soul|Neşe|Dans; Odak; Yolculuk|Ethio-Jazz; Mali Blues; Afrobeat; Scandinavian Jazz; Psychedelic Soul
Western Europe; Anatolia; Caucasus|Vokal/Şarkı; Gelenek|Melankoli|Dans|Kadın Polifonisi; 12 Dilde Repertuvar; Karadeniz–Anadolu–Sefarad–Oksitan
Northern Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Folk-Pop; Blues Tavırlı Vokal; Kopenhag; Bağımsız Kayıt
Southern Africa; West Africa; Western Europe; Fusion|Gelenek; Klasik/Modern Klasik|Epik|Dans|Çello + Batı Afrika Perküsyonu + Yaylı Kadrosu; Abel Selaocoe Projesi
North America|Blues|Melankoli|Gece|Cleveland Blues; Electric Blues; Guitar; Piano Boogie
North America|Gelenek|Neşe|Dans|Progressive Bluegrass; Mandolin/Keman/Gitar Üçlüsü
North America|Blues|Karanlık|Spor|Fingerstyle Blues + Punk; Los Angeles
North America|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Bluegrass; Americana; Folk; Country Roots
Caribbean|Vokal/Şarkı|Melankoli|Dans; Gece|Son Cubano; Trova; Cuarteto Geleneği
Latin America|Gelenek; Vokal/Şarkı|Melankoli|Gece|And Halk Müziği; Nueva Canción Öncüsü; Gitar
North America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Karanlık|Gece|Şarkı-Yazarı Düeti; Vokal Harmoni
Balkans|Vokal/Şarkı; Nefesli/Bando|Neşe|Dans; Odak|Makedon Roman Repertuvarı; Çalgi/Nefesli
Balkans|Gelenek; Vokal/Şarkı|Neşe|Dans; Gece; Yolculuk|Makedon Halk Şarkıları; İki Sesli Vokal; Geleneksel Repertuvar
Balkans|Vokal/Şarkı; Gelenek|Neşe|Dans; Yolculuk|Sevdalinka; Saz/Gitar; Saraybosna
Balkans; Western Europe; Fusion|Klasik/Modern Klasik; Elektronik|Melankoli|Odak|Çello; Sevdalinka; Bosna Etnosu; Elektronik; Çağdaş Klasik
Levant; Anatolia|Gelenek|Neşe|Dans|Kürtçe Repertuvar
Mediterranean Europe|Psikedelik/Rock|Melankoli|Gece|Sicilya Folk; Arap-Akdeniz Füzyon; Psikedelik Soul
Western Europe; Levant; West Africa; Fusion|Caz; Psikedelik/Rock|Karanlık|Odak; Yolculuk|Psych-Jazz; Mikrotonal Gitar; Ortadoğu/Balkan/Batı Afrika Groove'ları; Brüksel
Western Europe|Tel/Enstrümantal|Neşe|Dans; Spor|Hurda Enstrümanlar; Deneysel Folk; Enstrümantal
West Africa|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece|Balafon Ustası; Habib Koité Kadrosu; Bamako
East & Central Asia|Reggae/Dub; Gelenek|Neşe|Odak; Yolculuk|Wa Halk Müziği; Yunnan Reggae; Kadın Vokal Grubu
East Africa|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|Kenyan Contemporary Folk; Kalimba; Udu; African Popular Music
North America|Caz|Dinginlik|Yolculuk|Post-Bop; Serbest Caz; Perküsyonlu Kuintet
Northern Europe|Elektronik; Psikedelik/Rock|Neşe|Odak|Afro-Psych Groove; Kopenhag
Iberia|Vokal/Şarkı|Romantik|Dans|Çağdaş Fado; Minimal Prodüksiyon; İber İşbirliği
West Africa|Vokal/Şarkı; Blues|Melankoli|Gece|Dioula Vokali; Akustik Gitar; Bobo-Dioulasso
Western Europe|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans; Odak|House; French Touch; Minimal; Electronic Pop
North Africa; Western Europe|Klasik/Modern Klasik; Gelenek|Neşe|Odak; Gece|Mısır Şarkı Geleneği + Yaylı Düzenleme; Scotland
Eastern Europe|Gelenek|Melankoli|Dans; Yolculuk|Ethno-Chaos; Ukrayna Halk Müziği; Avant-Folk
North America; South Asia; East & Central Asia; Fusion|Caz|Melankoli|Yolculuk|Amerikana Füzyon; Hint Vokal; Caz Ritim
Western Europe|Gelenek|Neşe|Dans; Spor|Çok Dilli Folk; Chanson; Ada Müzikleri Füzyonu
North America|Tel/Enstrümantal; Caz|Neşe|Odak|Oud-Keman; Ortadoğu-Balkan Füzyon; Enstrümantal
Western Europe|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Gece; Yolculuk|Frizce Söz; Fado/Dünya Etkileri
Iberia; North America|Gelenek; Vokal/Şarkı|Melankoli|Odak; Gece; Yolculuk|Spanish Folk; Voice/Guitar; Traditional Song
Eastern Europe; South Asia; Fusion|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk|Spiritual Jazz; Hindustani Raga; Sarangi/Tabla
Anatolia|Psikedelik/Rock|Neşe|Odak; Yolculuk|Anadolu Psychedelia; Surf Rock; Enstrümantal
Caribbean|Vokal/Şarkı; Klasik/Modern Klasik|Romantik|Gece|Çello-Vokal; Latin Oda Müziği; Paris
Southeast Asia|Funk/Soul; Psikedelik/Rock|Neşe|Dans|Panturan Soul-Funk; Tarling; Terakota Enstrümanlar
West Africa|Vokal/Şarkı|Neşe|Dans; Odak; Yolculuk|Gana Halk Müziği; Highlife; Perküsyon-Vokal Düeti
Southeast Asia|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk; Spor|Khmer Pop; 1960'lar Kamboçya Rock; Ramvong
North America|Funk/Soul; Caz|Melankoli|Gece|Çoklu Enstrüman; Soul/Gospel; New York
Western Europe|Funk/Soul|Karanlık|Dans; Odak; Spor|Balkan Brass; Afrobeat; Enstrümantal Füzyon
Western Europe|Blues; Caz|Neşe|Dans; Odak|Boogie-Woogie Piyanosu; Stride
West Africa; Western Europe; Fusion|Reggae/Dub; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Afrika Reggae'si; Politik Şarkı; Fransız Şansonu; Batı Afrika Akustik Çalgıları
Western Europe; Fusion|Funk/Soul; Elektronik|Neşe|Dans; Spor|Soul; Latin; Afrobeat; Reggae; Acid Jazz; Trip-Hop; Psikedelya
Western Europe|Caz|Neşe|Yolculuk; Spor|Ortadoğu Cazı; Kamara Füzyon; Çok Kültürlü Topluluğu
East Africa|Vokal/Şarkı|Neşe|Dans|Somali Repertuvarı; Hargeisa; Savaş Hemşireliğinden Sahneye
West Africa|Psikedelik/Rock|Neşe|Yolculuk|Tuareg Rock; Assouf; Çöl Blues
Western Europe|Caz; Nefesli/Bando|Neşe|Dans; Spor|Caz-Punk/Swing; Dortmund
South Asia; North America; North Africa; Fusion|Blues|Melankoli|Yolculuk|Gazel; Pencap Folk; Sahra Blues Etkisi
Iranian World|Elektronik|Karanlık|Dans; Gece|Fars Pop; Elektronik; Diaspora Funk
Levant; North Africa; Fusion|Gelenek; Klasik/Modern Klasik|Epik|Dans; Odak|Arap-Endelüs Formları (touchia/samai); İbranice Özgün Parçalar; Kame'a
Levant|Vokal/Şarkı; Gelenek|Melankoli|Yolculuk|Arap Vokal Geleneği; Levant Repertuvarı
North Africa|Tel/Enstrümantal; Caz|Melankoli|Dans; Odak; Yolculuk; Spor|Ud Virtüozluğu; Tarab + Caz; Kahire
Mediterranean Europe; Iranian World; Fusion|Gelenek; Tel/Enstrümantal|Epik|Odak|Ortaçağ Lavtası/Ud + Afgan Rubabı
Levant|Elektronik; Tel/Enstrümantal|Neşe|Dans|Mandolin; Ortadoğu Telli Çalgılar; Tribal Elektronik
Western Europe|Vokal/Şarkı|Melankoli|Gece|Çok Dilli Folk; Kamera Pop; Vokal Üçlüsü
West Africa|Psikedelik/Rock|Melankoli|Odak; Yolculuk; Spor|Tuareg Rock; Çöl Blues; Psikedelik
Western Europe|Pop/Şarkı-Yazarı|Neşe|Dans; Odak|Chanson; Sefarad-Kuzey Afrika Etkisi
East Africa|Caz|Melankoli|Odak; Gece; Yolculuk|Ethio-Jazz; Saksafon; Addis Altın Çağ
Western Europe|Pop/Şarkı-Yazarı|Karanlık|Dans|Brecht Kabaresi; Karanlık Vodvil; Punk Kabare
Western Europe|Psikedelik/Rock; Elektronik|Karanlık|Dans; Yolculuk; Spor|Post-Punk + Asit Techno
Oceania|Vokal/Şarkı; Gelenek|Neşe|Odak|Aborjin Vokali; Numbulwar; Wubuy Dili
Iberia|Nefesli/Bando; Psikedelik/Rock|Karanlık|Odak; Yolculuk; Spor|Catalan Rock; Cobla; Sardana
North America|Blues|Karanlık|Yolculuk|Blues-Punk İkilisi; Org + Davul; Nashville
Balkans|Psikedelik/Rock|Epik|Odak; Yolculuk; Spor|Hırvat Enstrümantal Psych/Post-Rock; Međimurje Hattı
Eastern Europe; North Africa; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Macar Halk Müziği + Faslı Vokal
Southern Africa|Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|Zamrock; 70'ler Garaj/Fuzz
Western Europe|Caz; Funk/Soul|Neşe|Odak; Yolculuk; Spor|British Nu Jazz; Afro-Rock; Jazz-Funk; Groove
Iranian World|Elektronik|Neşe|Dans|Fars Pop; Elektronik; Diaspora Funk
Latin America|Gelenek; Elektronik|Dinginlik|Odak; Gece|Latin Folk; Electrocumbia
Levant|Psikedelik/Rock; Funk/Soul|Neşe|Dans; Odak|Anadolu Psych; Saz; Dabke; Bandari; Dub; Batov Records
Eastern Europe|Gelenek; Nefesli/Bando|Neşe|Dans; Spor|Romen Lăutari Orkestra Geleneği
Latin America|Elektronik|Neşe|Dans; Odak|Latin House/Tropikal Elektronik (başlıktan çıkarım)
South Asia|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Newa Müziği; Nepal Yerli Çalgıları; Çağdaş Folk; Tamang Selo Etkisi
Caribbean; Western Europe; Fusion|Elektronik; Reggae/Dub|Neşe|Dans|Dub; Digital Production; Studio Fusion
West Africa|Tel/Enstrümantal; Gelenek|Dinginlik|Odak; Yolculuk|Akustik Gitar İkilisi; Batı Afrika Griot Repertuvarı; 12 Telli ve 6 Telli Gitar
Iranian World; Mediterranean Europe; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Gece; Yolculuk|Ney + Rubab + Gitar; Roots Revival Projesi
Anatolia|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|Organic House; Electronic Indie; Turkish Vocal; Acoustic-Electronic
Levant|Klasik/Modern Klasik; Caz|Dinginlik|Dans; Odak; Yolculuk|Oryantal Füzyon Piyano; Tango, Flamenko, Latin ve Caz Etkileri
Levant|Gelenek|Karanlık|Yolculuk; Spor|Deneysel Arap Müziği; Serbest Doğaçlama; Noise-Folk
Iranian World|Caz; Gelenek|Dinginlik|Odak; Spor|Mikrotonal Caz; Film Farsi Estetiği; Tahran Sahnesi
Western Europe|Nefesli/Bando; Caz|Dinginlik|Odak; Gece|Klarnet Üçlüsü; Klezmer Köklü Oda Cazı
Oceania|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Hint Etkili Enstrümantal Psych; Melbourne; Anonim Kimlik
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Dans|Indie Folk-Pop; Fingerstyle Akustik; Normandiya; Wales
Oceania; Western Europe|Tel/Enstrümantal; Gelenek|Melankoli|Odak; Gece|Çok Tarzlı Çello; Folk/Bluegrass + Batı-Dışı Ses Dünyaları
Western Europe; North Africa; Fusion|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk|Avrupa Oda Cazı + Mısır Ud/Nay; Cairo Steps
East & Central Asia|Elektronik|Dinginlik|Odak; Yolculuk|Caz; Film Müziği; Latin-Elektronik Füzyon
East & Central Asia|Elektronik|Melankoli|Yolculuk|Caz; Film Müziği; Latin-Elektronik Füzyon
East Africa|Gelenek; Psikedelik/Rock|Neşe|Dans; Spor|Maloya; Kreol Söz; Çağdaş Prodüksiyon
Balkans|Klasik/Modern Klasik|Melankoli|Odak; Gece|Film Müziği; Modern Klasik; Yunan Kamara Geleneği
Western Europe|Klasik/Modern Klasik; Elektronik|Neşe|Dans|Film Müziği; Çağdaş Fransız Sineması
North America|Gelenek; Blues|Melankoli|Gece|Siyah Telli Çalgı Geleneği; Banjo/Keman; Pulitzer
Iberia|Vokal/Şarkı; Gelenek|Melankoli|Gece|Flamenko Cantesi; Genç Nesil
Southeast Asia|Gelenek; Reggae/Dub|Neşe|Dans; Odak; Yolculuk|Molam; Phin; Khaen; Dub; Thai Funk
South Asia|Vokal/Şarkı|Epik|Gece; Yolculuk|Kavvali; Sufi İlahi Geleneği; Pencap Vokali
Balkans|Gelenek; Vokal/Şarkı|Melankoli|Gece|Berat Urban Folk; Albanian Traditional
Western Europe|Pop/Şarkı-Yazarı; Elektronik|Melankoli|Odak; Gece; Yolculuk|Karanlık Folk; Third Eye Foundation; Bristol
Iberia|Klasik/Modern Klasik|Dinginlik|Dans|Polifonik Vokal Düeti; Akdeniz İlahileri; Erken Müzik; Catalonia
Mediterranean Europe|Vokal/Şarkı; Gelenek|Melankoli|Dans; Odak; Spor|Güney İtalya Halk Vokali; Ağıt Repertuvarı
Anatolia; Caucasus|Vokal/Şarkı; Gelenek|Romantik|Gece|Kürtçe-Türkçe Repertuvar; Kafkas Vokali
Balkans|Vokal/Şarkı|Melankoli|Gece|Yunan Entekno; Şiirsel Şarkı; Sanat Müziği
Eastern Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Solo Akustik Gitar; Sokak/Gezgin İcra; Flamenko-Balkan Etkileri
Levant|Tel/Enstrümantal; Caz|Epik|Dans; Odak|Solo Kontrbas; Perküsif Teknik; Loop
North Africa|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Tunus Vokali; Çağdaş Prodüksiyon
Mediterranean Europe|Vokal/Şarkı; Gelenek|Melankoli|Odak; Yolculuk|Güney İtalya Halk Vokali; Ağıt Repertuvarı
Iberia|Nefesli/Bando|Romantik|Gece|Balkan Brass; Roman Müziği; Klezmer Füzyon
North America|Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece|Handpan/Hang; Solo Enstrümantal; Philadelphia
Iberia; Western Europe; Fusion|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Yolculuk|Portekiz Halk Müziği; Adufe; Caz; Psikedelya; Elektronik
Northern Europe|Pop/Şarkı-Yazarı|Dinginlik|Gece; Yolculuk|İzlanda Indie Folk; Viola-Charango; Minimal Şarkı
North Africa|Vokal/Şarkı|Neşe|Dans|Cezayir/Arap Repertuvarı; Dans-Vokal; Kahire
Levant|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|Hebrew Song; Folk; Live Looping; Falafel Pop
Iberia|Gelenek|Neşe|Dans|Galiçya Halk Müziği; Xota; Pandeireta Geleneği
Iberia|Gelenek|Melankoli|Dans|Galiçya Folk; Keman-Akordeon Düeti; Kabare Pop
Caucasus|Gelenek|Romantik|Dans; Yolculuk; Spor|Gürcü Halk Müziği; Chonguri-Panduri; Çok Sesli Vokal
Iberia|Vokal/Şarkı|Melankoli|Gece|Çağdaş Fado; Şiir Uyarlamaları; Lizbon
Balkans|Elektronik|Melankoli|Dans; Spor|Yunan Retro Sample; Laiko; Downtempo
Iranian World|Vokal/Şarkı; Gelenek|Melankoli|Dans; Yolculuk|Fars Klasik/Halk Repertuvarı; Diaspora; Canlı Kayıt
Eastern Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Polonya Indie Pop; Şarkı Yazımı; Akustik Rock
West Africa|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Bamana Repertuvarı; Ngoni + Elektrikli Gitar; Bağımsız Söz
Latin America|Vokal/Şarkı; Caz|Neşe|Dans; Yolculuk; Spor|Bolero; Son; Swing; Latin Jazz; Orchestrated Song
Western Europe|Klasik/Modern Klasik; Pop/Şarkı-Yazarı|Karanlık|Dans|Karanlık Art-Pop; Piyano; Viyana
Anatolia|Vokal/Şarkı|Romantik|Gece|Kürtçe Şarkı; Anadolu Vokal Geleneği
Anatolia|Pop/Şarkı-Yazarı|Neşe|Dans; Odak; Spor|Oyuncu/Komedyen Kökenli Müzik Projesi; İstanbul
Western Europe|Tel/Enstrümantal|Melankoli|Dans; Odak|Flamenko Gitar İkilisi; Alman Kadro
Western Europe; Mediterranean Europe|Vokal/Şarkı; Elektronik|Melankoli|Dans|Loop-Based Song; Cello; Guitar; Drum; Improvised Voice; Ceremonial Music
West Africa; North Africa; Fusion|Vokal/Şarkı; Gelenek|Neşe|Dans; Spor|7 Dilde Şarkı Sözü Yazarlığı; Bambara; Soul + Saudade
Western Europe|Tel/Enstrümantal; Caz|Karanlık|Odak|Gitar; Arap-Caz Füzyonu; Düzenleme; Paris
Iranian World; East & Central Asia; Western Europe; Fusion|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Yolculuk|Persian Poetry; Choir/Orchestra; Rumi
Northern Europe; Levant; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk|Avangart Pop + Suriye Dabkesi
Southern Africa|Psikedelik/Rock|Dinginlik|Yolculuk|Zamrock; 70'ler Garaj/Fuzz
North Africa; West Africa|Blues|Neşe|Dans; Yolculuk; Spor|Tuareg Assouf; Çöl Blues; Tamanrasset Sahnesi
Anatolia|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Dersim/Zaza Repertuvarı; Alan Kayıtları; Kahraman Kardeşler
Oceania; Balkans; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak; Gece; Yolculuk|Ud-Laouto; Balkan/Akdeniz Cazı; Brisbane
Western Europe|Caz; Pop/Şarkı-Yazarı|Melankoli|Dans; Yolculuk|Londra Sahnesi; Afro-Latin Etkili İndie
Western Europe; Levant|Caz|Neşe|Dans; Odak; Spor|Londra Cazı; Orta Doğu Duyarlılığı; Ud + Nefesliler; Yedili
Latin America|Klasik/Modern Klasik; Pop/Şarkı-Yazarı|Melankoli|Odak; Gece|Film Müziği; Ronroco; Rock Nacional; Bajofondo
Iberia; Levant; Mediterranean Europe; Fusion|Caz; Tel/Enstrümantal|Karanlık|Odak; Yolculuk|Experimental Jazz; Mediterranean Modal Music
Mediterranean Europe|Psikedelik/Rock|Melankoli|Gece|İtalyan Cantautore; Sicilya Folk; Rock
Western Europe|Elektronik|Melankoli|Odak|Klezmer Füzyon; Balkan; Elektronik Folk
Levant|Klasik/Modern Klasik|Melankoli|Gece|Film Müziği; Nadine Labaki Filmleri; Beyrut
Western Europe|Elektronik|Neşe|Dans|Indie Pop; Deneysel Folk; Elektronik
North America|Psikedelik/Rock|Melankoli|Dans|Amerikana; Noir Folk; Rockabilly-Blues
Levant|Caz|Neşe|Dans; Odak; Yolculuk|Piyano; Arap Cazı; Şarkı Yazımı
Balkans|Psikedelik/Rock; Funk/Soul|Neşe|Dans; Odak|Anadolu Rock Etkisi; Tay-Kamboçya Funk; Enstrümantal Psikedelia
Western Europe; Balkans; Anatolia; Fusion|Nefesli/Bando|Neşe|Dans; Spor|Balkan/Türk Roman Nefesli Geleneği; Paris
Levant|Psikedelik/Rock|Melankoli|Dans; Gece|Psychedelic Rock; Akdeniz Surf; Doğu Rock
North Africa|Elektronik|Melankoli|Dans; Yolculuk; Spor|Arabic Fusion; Trip-Hop; Orkestral Arap Pop
Western Europe; Eastern Europe; North Africa; Fusion|Klasik/Modern Klasik|Melankoli|Yolculuk|Keman; Klezmer-Kamara; Arap Vokal Konuğu
West Africa|Elektronik; Gelenek|Neşe|Dans; Spor|Vodun Ritimleri + Elektronik; Kadın Topluluğu; Lomé
North America|Elektronik; Vokal/Şarkı|Neşe|Dans|New York Global Beat; Turntables on the Hudson
Western Europe; South Asia; Fusion|Caz; Gelenek|Karanlık|Dans; Odak; Spor|World Jazz; Indian Sitar; Trumpet; Double Bass; Percussion; Asian Vibe
Eastern Europe|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece|Tango Nuevo; Bandoneón Kuarteti
Northern Europe|Gelenek|Neşe|Dans; Odak; Spor|Fin Akordeonu; Solo; Nefes ve Ritim
Balkans|Klasik/Modern Klasik; Pop/Şarkı-Yazarı|Melankoli|Odak; Gece|Yunan Sanat Şarkısı; Film Müziği; Buzuki
Mediterranean Europe|Klasik/Modern Klasik; Tel/Enstrümantal|Romantik|Odak; Gece|İki Çello Düosu; Barok–Caz–Rock Düzenlemeleri; Bologna
Anatolia|Vokal/Şarkı; Gelenek|Neşe|Dans; Spor|Karadeniz/Laz Repertuvarı; Kemençe Hattı
Levant|Psikedelik/Rock; Funk/Soul|Neşe|Dans; Odak|Anadolu Psych; Levanten Groove; Batov Records
East & Central Asia|Elektronik; Klasik/Modern Klasik|Dinginlik|Dans; Odak|Video Game Soundtrack; Cinematic Electronic
Levant|Pop/Şarkı-Yazarı; Elektronik|Neşe|Yolculuk|Alternative Pop; Spoken Vocal; Guest Ensemble
Balkans|Elektronik; Hip-Hop/Rap|Neşe|Dans; Spor|Yunan Elektronik/Downtempo
Fusion|Caz; Gelenek|Neşe|Dans; Yolculuk; Spor|Doğu-Batı Füzyonu; Caz; Hicaz/Kürdi Makam Referansları; Doğaçlama
Balkans|Tel/Enstrümantal; Pop/Şarkı-Yazarı|Romantik|Odak; Gece|Balkan Gitarı; Eklektik Füzyon; Paris
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|80'ler-90'lar İngiliz Şarkı-Yazarı; Alto Vokal
Anatolia|Tel/Enstrümantal|Melankoli|Odak; Gece|Mikrotonal Gitar; Anadolu Makamı; Enstrümantal Düet
North Africa|Caz|Melankoli|Odak; Gece|Oud; Arap Klasik Müziği; Kamara Caz
Anatolia; Western Europe|Gelenek; Tel/Enstrümantal|Epik|Dans; Odak; Yolculuk; Spor|Anadolu Telleri; Avrupa Kadrosuyla Kayıt
Western Europe; Latin America; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece|Enstrümantal Çöl Gitarı; İki Kardeş; Easy Eye Sound
Northern Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece|Enstrümantal Gitar; Latin Füzyon; Downtempo
North America|Tel/Enstrümantal|Dinginlik|Odak; Gece|Nuevo Flamenko; Rumba Gitarı; Toronto
Iberia|Tel/Enstrümantal|Neşe|Odak|Flamenko Gitarı; Córdoba Ekolü
North Africa; Western Europe|Caz; Vokal/Şarkı|Melankoli|Gece|Amaziğ/Fas Vokali + Caz; Berber Taxi
Western Europe; Anatolia|Elektronik|Neşe|Dans; Odak; Yolculuk|Viyana Lounge/Downtempo; Kürt Kökenli Prodüktör; Hôtel Costes Çevresi
Western Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece|Enstrümantal; Dans ve Tiyatro Müziği; Minimal Folk
Northern Europe|Gelenek; Klasik/Modern Klasik|Melankoli|Odak|Estonya Sahne Müziği; Yerel Halk Öğeleri
North America|Psikedelik/Rock|Neşe|Dans; Odak; Yolculuk|Prog Rock Üçlüsü; Enstrümantal
Anatolia|Elektronik|Dinginlik|Dans; Odak; Yolculuk; Spor|Elektronik; Anadolu Psychedelia; Organik House
North America|Funk/Soul|Melankoli|Odak|New Orleans R&B; Piyano; Besteci-Prodüktör
Fusion|Psikedelik/Rock; Pop/Şarkı-Yazarı|Karanlık|Odak; Yolculuk; Spor|Yapay Zekâ Destekli AOR; Arena Rock; Progresif Rock; Metal
Levant; Mediterranean Europe; Fusion|Klasik/Modern Klasik|Romantik|Gece|Mandolin; Akdeniz Halk Repertuvarı; Kamara Düzenleme
North Africa|Psikedelik/Rock|Melankoli|Dans|Tuareg Rock; Çöl Blues; Assouf
Western Europe; South Asia; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Yolculuk|İngiliz Folkü + Hint Klasiği; Sitar-Keman
West Africa|Funk/Soul; Gelenek|Neşe|Yolculuk; Spor|Gine Afro-Funk; 70'ler; Analog Africa Yeniden Basımı
West Africa|Gelenek; Funk/Soul|Neşe|Dans; Yolculuk; Spor|Assiko; Wolof Vokal; Dakar/Grand Yoff
Western Europe; Iberia|Elektronik; Klasik/Modern Klasik|Dinginlik|Odak|Musique Concrète; Electroacoustic; Avant-Garde
West Africa|Psikedelik/Rock|Melankoli|Yolculuk|Tuareg Rock; Assouf; Çöl Blues
Latin America|Elektronik|Dinginlik|Odak; Gece|Deneysel Elektronik; Other People; New York-Santiago
North Africa|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Tuareg Rock; Çöl Blues; Assouf
North America; West Africa; Fusion|Caz; Gelenek|Dinginlik|Odak|World Jazz; Banjo; Kora
Levant; North America|Elektronik; Funk/Soul|Neşe|Dans|Global Funk; Ambient; Electro; Oud; Ney; Saz; Instrumental
East Africa; Western Europe; Fusion|Elektronik|Neşe|Dans; Odak; Yolculuk|Buganda Davulu + Techno; Kampala
Eastern Europe|Gelenek|Melankoli|Gece|Ethno-Chaos; Ukrayna Halk Müziği; Avant-Folk
South Asia|Tel/Enstrümantal; Klasik/Modern Klasik|Karanlık|Dans; Odak|Sitar + Sarod + Perküsyon; Hindustani Çağdaşı
North Africa|Vokal/Şarkı|Epik|Dans; Spor|Sahrawi Repertuvarı; Sürgün/Direniş Şarkıları
West Africa|Blues; Psikedelik/Rock|Epik|Dans|Tuareg Çöl Blues'unun Kurucuları; Tamashek; Assouf
Northern Europe|Psikedelik/Rock|Karanlık|Odak; Yolculuk|Psikedelik Rock; Afro-Krautrock; Maskeli Kolektif
Caribbean|Vokal/Şarkı; Klasik/Modern Klasik|Neşe|Dans; Yolculuk|Çello-Vokal; Latin Oda Müziği; Paris
Caribbean; North America|Reggae/Dub|Neşe|Dans|Analog Dub; Tek Kişilik Prodüksiyon; Los Angeles
Latin America|Psikedelik/Rock; Vokal/Şarkı|Neşe|Dans|Brezilya Deneysel Pop'u; Tropikalya Mirası; Rio
Caribbean|Funk/Soul; Gelenek|Neşe|Dans; Odak|Guyana Folk-Funk; Kwe-Kwe/Afro-Karayip Perküsyon
Mediterranean Europe|Gelenek|Neşe|Gece|Salento Pizzica/Tarantella; Kadın Vokal; Puglia
Caucasus|Gelenek|Neşe|Dans|Ermeni Halk Müziği; Balkan Füzyon; Pop
Western Europe; West Africa; Fusion|Pop/Şarkı-Yazarı; Nefesli/Bando|Neşe|Dans; Yolculuk|French Chanson; West African Brass; Collective Protest Song
Anatolia|Caz|Dinginlik|Odak; Gece; Yolculuk|Contemporary Jazz; Piano/Guitar Quartet; Improvisation
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Gece; Yolculuk|İskoç ve İrlanda Halk Müziği; Şarkıcı-Şarkı Yazarı; Folk-Caz; Country; Glasgow/Fife
North America; East Africa; Fusion|Caz|Neşe|Yolculuk; Spor|Big Band; Ethio-Jazz Yorumları; Boston
North America|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk|Saksofon + Perküsyon; Yaqui Kökeni; Daptone Çevresi
North Africa|Pop/Şarkı-Yazarı|Melankoli|Dans|Mısır Protest Şarkısı; Tahrir; Sürgün
Levant; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Gece|Ud-Kanun-Ney + Avrupa Perküsyonu; Diaspora Topluluğu
North Africa; Levant|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Alternative Arabic Music; Experimental Electronics; Oud; Buzuq
Eastern Europe; Anatolia; Fusion|Elektronik; Gelenek|Melankoli|Dans; Yolculuk|Türk Halk Ezgileri + Elektronik Prodüksiyon
Levant|Caz|Dinginlik|Odak|Saksofon; Arap Ezgileri + Caz; Paris
North America|Klasik/Modern Klasik|Melankoli|Odak; Gece|Deneysel Klasik; Perküsyon; Kanon-Round Formu
North Africa|Elektronik; Tel/Enstrümantal|Epik|Odak|Ud + Elektronik; Paris
Northern Europe|Caz|Neşe|Odak; Spor|Enstrümantal; Kütüphane Müziği; Anadolu-Japon Etkili Caz-Funk
Anatolia; Western Europe|Tel/Enstrümantal; Caz|Dinginlik|Odak; Yolculuk|Perdesiz Ud; Anadolu Makamı + Caz; Amsterdam
North America|Caz; Funk/Soul|Neşe|Odak; Yolculuk|Spiritual Jazz; Afrobeat; Free Improvisation
Oceania|Reggae/Dub; Funk/Soul|Neşe|Dans; Odak; Yolculuk; Spor|Dub-Soul Kolektifi; Doğaçlama; Wellington
East & Central Asia; Northern Europe; Fusion|Elektronik; Reggae/Dub|Neşe|Yolculuk; Spor|Dub; Krautrock; Chinese Instrumental
East & Central Asia|Gelenek; Vokal/Şarkı|Melankoli|Gece|Tuva Gırtlak Şarkıcılığı; Kargyraa Ustası
Southeast Asia|Gelenek; Tel/Enstrümantal|Neşe|Yolculuk|Endonezya Etnik Füzyon; Çoklu Enstrüman; Ses İyileştirme; Ubud/Bali
Northern Europe|Caz|Karanlık|Spor|Org Üçlüsü; Caz-Rock; Kopenhag
Northern Europe; East Africa; Western Europe|Psikedelik/Rock; Hip-Hop/Rap|Neşe|Dans; Odak; Yolculuk; Spor|Afro-Krautrock; Multilingual Rap; Remix
Anatolia|Elektronik; Caz|Karanlık|Dans; Odak; Yolculuk|Caz; Elektronik; Nublu Sahnesi
North Africa|Reggae/Dub|Neşe|Yolculuk|Libya Reggae'si; 80'ler; Habibi Funk
Levant|Psikedelik/Rock; Funk/Soul|Neşe|Dans; Odak|Anadolu Rock/Space Funk; Bourekas Estetiği; French House; Kame'a
Iberia; Western Europe; Levant; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Gece|Murcian Folk; Mandolin Family; Cello; Arabic-Andalusian; Ney; Oud; Qanun
Caribbean; West Africa; Fusion|Caz; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Piyano + Kora; Afro-Küba – Mande Buluşması
North America; West Africa; Fusion|Caz; Gelenek|Dinginlik|Dans|Caz Vokali + Yeşil Burun Gitarı/Cavaquinho
Latin America|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Pan-Latin Amerikan Repertuvar; Huayño/Cumbia/Taquirari; Quena; La Paz
North America|Gelenek|Melankoli|Gece|Old-Time/Bluegrass; Kadın İkili Geleneği
Anatolia|Gelenek|Dinginlik|Odak; Yolculuk; Spor|Sufi Müziği; Ney/Zurna; Ortadoğu Füzyonu
East & Central Asia|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans; Yolculuk|Chinese Indie Pop; Dance Pop; Bedroom Production
Eastern Europe|Gelenek|Melankoli|Yolculuk|Ethno-Chaos; Ukrayna Halk Müziği; Avant-Folk
Oceania; South Asia; Fusion|Caz; Gelenek|Neşe|Odak; Yolculuk|Indian Classical; Tabla; Jazz Fusion; Blues; Didjeribone
Anatolia|Psikedelik/Rock|Melankoli|Yolculuk; Spor|Anadolu Psychedelia; Modern Düzenleme; Kadın Vokal
North America|Caz; Tel/Enstrümantal|Karanlık|Odak; Yolculuk|Avangard Arp; Spiritual Jazz; New Orleans
West Africa|Blues|Dinginlik|Dans; Odak; Yolculuk|Kamele Ngoni; Mande Müziği; Çöl Blues
West Africa|Blues; Psikedelik/Rock|Neşe|Yolculuk|Tuareg Çöl Blues'unun Kurucuları; Tamashek; Assouf
Western Europe; South Asia; North Africa; Fusion|Tel/Enstrümantal|Epik|Dans; Odak; Spor|Buzuk/Ud/Gitar; Roman-Hint-Mağrip Hattı
North America|Gelenek; Blues|Romantik|Dans; Yolculuk|Cajun/Louisiana + Mississippi Fife and Drum
Western Europe; Iberia; Southern Africa; Fusion|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans; Gece; Yolculuk|Art Pop; Electronica; Saudade; Alternative Latin; Afro-Pop
Western Europe|Psikedelik/Rock|Dinginlik|Dans; Odak; Spor|Ethio-Jazz; Balkan-Doğu Füzyon; Psikedelik Enstrümantal
Western Europe|Tel/Enstrümantal; Gelenek|Melankoli|Dans; Yolculuk|Arp/Keman; Kelt Repertuvarı; Wales
Mediterranean Europe|Gelenek; Caz|Neşe|Odak|Diyatonik Akordeon + Klarnet; Toskana Hattı
Northern Europe; Anatolia|Gelenek; Vokal/Şarkı|Neşe|Dans|Anadolu/Kürt Repertuvarı + Nordik Düzenleme; Kopenhag
Western Europe; Balkans; Anatolia; Levant; Fusion|Psikedelik/Rock; Vokal/Şarkı|Neşe|Spor|Buzuki + Hammond Org; Rebetiko/Tsifteteli; Berlin
Western Europe|Elektronik|Neşe|Dans; Odak; Yolculuk; Spor|Downtempo/Bossa; Lounge; Hannover
North America|Blues; Caz|Melankoli|Odak|Kornet/Gitar; Mississippi Kökeni; Afrika Bağlantıları
West Africa|Blues|Romantik|Yolculuk|Songhai Blues; Elektrikli Gitar; Niafunké Ekolü
Iranian World|Caz|Melankoli|Dans; Spor|Afgan Halk Müziği; Deneysel; Politik Caz
South Asia; North America; Fusion|Tel/Enstrümantal; Klasik/Modern Klasik|Epik|Odak; Yolculuk|Sarod + Sitar + Batı Yaylı Dörtlüsü + Tabla
Anatolia|Elektronik|Neşe|Odak; Yolculuk; Spor|Elektronik; Anadolu Psychedelia; Organik House
West Africa|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Kora; Griot Geleneği; Kora Pop'unun Öncüsü
Northern Europe; East Africa|Funk/Soul; Pop/Şarkı-Yazarı|Romantik|Dans|İsveç-Etiyopya Kökeni; Soul-Folk; Göteborg
North America|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Gece|Québec Folk; Şiir Bestelemeleri; Montreal
Latin America|Vokal/Şarkı; Gelenek|Melankoli|Dans; Yolculuk|Şili Bolero/Cebolla + Tango Vokali
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Dans; Odak; Spor|Latin Alternatif; Bolero/Ranchera Etkisi; Meksika Yerleşik
Western Europe|Caz; Pop/Şarkı-Yazarı|Dinginlik|Odak; Yolculuk|Halk Müziği; Caz; Ambient; Deneysel Şarkı Yazımı
Iberia; Fusion|Elektronik; Vokal/Şarkı|Neşe|Dans; Spor|Galician Folktronica; Asturian Folk; Ritual Pop
Western Europe|Gelenek|Neşe|Dans; Odak|Balkan; Klezmer; Roman Müziği
Western Europe|Gelenek|Dinginlik|Odak; Gece; Yolculuk|Balkan; Klezmer; Roman Müziği
Levant|Elektronik; Psikedelik/Rock|Karanlık|Dans; Odak; Yolculuk|Lübnan Avangardı; Omar Khorshid Yorumu; Darbuka + Elektronik
Iranian World; North America|Gelenek; Klasik/Modern Klasik|Neşe|Dans; Yolculuk|Tombak; Persian Rhythmic Cycles; Early Music; Contemporary Composition
Oceania; Eastern Europe; Fusion|Pop/Şarkı-Yazarı; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Akustik Dünya Müziği; El Yapımı Çalgılar; Flamenko/Son Küba Etkisi
Western Europe|Elektronik; Caz|Dinginlik|Odak|Sinematik Elektronika; Hint Sample'ları; Sitar; Beat-Making
Iberia|Vokal/Şarkı|Melankoli|Gece|Fado; Lizbon
Iberia|Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Portekiz Noir Gitarı; Enstrümantal İkili
Iberia|Gelenek; Vokal/Şarkı|Neşe|Dans|Baixo Douro Folkloru; Geleneksel Toplulukla Kayıt
Latin America|Vokal/Şarkı|Romantik|Gece|Slowcore Chanson; Meksika Bolero; Alternatif Folk
West Africa|Tel/Enstrümantal; Gelenek|Dinginlik|Dans; Odak; Gece|Çöl Bluesu Gitarı + Kora; Mande Geleneği
Northern Europe; North America; Fusion|Elektronik; Klasik/Modern Klasik|Karanlık|Odak; Yolculuk|Genişletilmiş Akordeon; Sampling; Elektronik; Yaylı Dörtlüsü; Çağdaş Oda Müziği
Levant|Elektronik|Epik|Dans; Yolculuk; Spor|Shamstep; Dabke; Elektro-Arap
Latin America; Western Europe|Gelenek; Klasik/Modern Klasik|Neşe|Dans; Yolculuk|Kolombiya Repertuvarı + Klasik Düzenleme; Salsa Efsanesi Konuk
Latin America; North America; Western Europe; Caribbean; Fusion|Caz|Dinginlik|Odak; Yolculuk|Afro-Brazilian Jazz; Avant-Jazz; Flute/Piano Improvisation
Northern Europe|Gelenek|Neşe|Yolculuk; Spor|Nordik Halk Müziği; Geniş Kadro
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Yolculuk|İngiliz Folk İkilisi; Suffolk; Çoklu Enstrüman
Western Europe|Gelenek; Psikedelik/Rock|Melankoli|Gece; Yolculuk|Karanlık Folk; İrlanda Balıadları; Lankum Çevresi
Iranian World|Gelenek; Tel/Enstrümantal|Romantik|Dans; Spor|Afgan Rubab/Ghichak Geleneği; Khushnawaz Hattı
Iberia; East & Central Asia; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Oud; Qobyz; Silk Road
East & Central Asia|Elektronik|Karanlık|Dans; Odak; Spor|Caz; Film Müziği; Latin-Elektronik Füzyon
Caribbean; North America|Vokal/Şarkı; Caz|Neşe|Dans; Yolculuk; Spor|Afro-Kuba + Caz; Kadın İkili; Toronto
East & Central Asia|Caz; Klasik/Modern Klasik|Karanlık|Odak|Avant-Garde Jazz; Dark Jazz; Modern Classical; Nu Jazz; Instrumental
Iranian World|Vokal/Şarkı|Melankoli|Dans; Yolculuk|İran Bölgesel Halk Müziği; Çok Dilli Vokal
Caribbean; Western Europe; Fusion|Caz; Tel/Enstrümantal|Dinginlik|Odak; Gece|Afro-Küba/Latin Cazı; Piyano; Çello; Oda Müziği Doğaçlaması
Levant|Psikedelik/Rock; Funk/Soul|Neşe|Odak; Spor|Ortadoğu Sörf/Funk; Batov Records; Tel Aviv
West Africa|Tel/Enstrümantal; Vokal/Şarkı|Dinginlik|Gece|Kora; Griot Geleneği; Saint-Louis
North America|Vokal/Şarkı|Neşe|Dans; Odak; Spor|Bolero; Cumbia; Enstrümantal Gitar Üçlüsü
East & Central Asia|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Pansori + Alternatif Pop; Seul
Western Europe; Levant; Fusion|Tel/Enstrümantal|Romantik|Dans; Odak; Yolculuk; Spor|Flamenko Gitar; Arap Perküsyonu; Doğu-Batı Düeti
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Enstrümantal Psikedelik; Uzakdoğu Etkili Funk; Kraut-Disco
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Caz; Spiritual Jazz; Elektronik Groove
Western Europe|Elektronik; Caz|Neşe|Dans; Odak|Çello + Davul İkilisi; Doğaçlama Elektronik; Paris
Western Europe|Caz|Karanlık|Odak; Spor|Tamamen Doğaçlama; Londra; Perküsyon Ağırlıklı
Western Europe; East & Central Asia; South Asia; Fusion|Klasik/Modern Klasik|Dinginlik|Gece|Gitar; Çok Kültürlü Yaylı Topluluğu; Film Müziği
East Africa|Caz|Dinginlik|Dans|Ethio-Jazz'ın Kurucusu; Vibrafon; Addis Ababa
Levant|Psikedelik/Rock|Neşe|Dans|Doğu Surf; Psikedelik Rock; Enstrümantal
North America; Western Europe|Caz; Elektronik|Melankoli|Dans; Gece|Caz Vokali + İngiliz Elektronik Prodüksiyonu
West Africa; North America; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Dans|Senegal Gitarı + Amerikan Parmak Stili
Western Europe; West Africa; Fusion|Tel/Enstrümantal; Blues|Neşe|Dans; Yolculuk|Çöl Bluesu Gitarı + Ritti (Tek Telli Keman)
Caribbean|Hip-Hop/Rap; Funk/Soul|Karanlık|Dans|Flüt + Rap; Cumbia/Dancehall; Paris
North America|Funk/Soul; Vokal/Şarkı|Neşe|Dans; Spor|Southern Soul; Gospel; New Orleans Brass
Latin America|Gelenek; Nefesli/Bando|Neşe|Dans; Odak|Mariachi; Los Angeles; Nati Cano Ekolü
Mediterranean Europe|Gelenek|Neşe|Yolculuk|Organetto; Tarantella/Pizzica; Güney İtalya
Balkans|Gelenek|Dinginlik|Odak; Gece|Balkan Akordeon + Keman İkilisi; Norveç Yerleşik
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Spor|İngiliz Folk; Parmak Stili Gitar; Birmingham
Western Europe|Caz; Gelenek|Neşe|Dans|Manouche/Gypsy Jazz; 1930'lar Paris; Klezmer-Balkan-Kelt Esintileri
Balkans; Mediterranean Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Odak|Greek Folk; Accordion; Cello
Eastern Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece|Hungarian Folk; Transylvanian Folk; Multilingual Song; Koboz; Vocal Storytelling
Iberia|Gelenek; Nefesli/Bando|Neşe|Dans|Mallorcan Folk; Flabiol; Xeremies; Folk-Rock; Mediterranean Roots
Levant|Caz|Dinginlik|Odak; Gece|Free Jazz; Gnawa; Spiritual Jazz; Gimbri
Anatolia; Western Europe; Fusion|Elektronik; Vokal/Şarkı|Melankoli|Dans; Yolculuk|Kürt Sufi Müziği; Saha Kayıtları; Setar; Elektronik; Irak Kürdistanı Ses Peyzajı
Western Europe; Iranian World; Fusion|Caz; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Persian Classical Vocal; Contemporary Jazz
South Asia|Vokal/Şarkı|Melankoli|Dans; Gece|Hint Vokal + Minimal Prodüksiyon; Hindustani/Urdu Söz
North Africa; Western Europe; Fusion|Funk/Soul; Gelenek|Neşe|Spor|Gnawa Funk; Chaâbi/Beldi Ritimleri; Afrobeat; Oriental Synth
Levant|Caz; Tel/Enstrümantal|Neşe|Dans; Yolculuk; Spor|Kanun; Irak Makamı + Caz; Brüksel
Eastern Europe; West Africa; Fusion|Gelenek; Psikedelik/Rock|Neşe|Yolculuk|Wolof Vocals; Kora; Trance; Experimental Guitar
Mediterranean Europe; Iranian World; Fusion|Gelenek; Elektronik|Karanlık|Yolculuk|Akdeniz Halk Müziği; İran Etkileri; Vokal/Perküsyon Gelenekleri; Elektronik Trans
Iberia; Western Europe|Gelenek; Klasik/Modern Klasik|Dinginlik|Odak; Gece|Flamenco; Contemporary Classical
Caribbean|Reggae/Dub|Karanlık|Dans; Yolculuk|Dub'ın Kurucusu; Black Ark Stüdyosu; Kingston
Western Europe; Anatolia|Vokal/Şarkı|Neşe|Gece|Anatolian Song; Turkish Poetry; Accordion; Clarinet; Acoustic Folk
West Africa; North America; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Dans; Gece; Yolculuk|West African Folk; Appalachian Banjo; Accordion
North Africa; Western Europe; Iberia; Fusion|Caz; Gelenek|Dinginlik|Odak; Yolculuk|Stambeli; Contemporary Jazz; North African Modal Music
West Africa|Blues|Neşe|Dans|Mali Blues; Afro-Pop; Gitar Düeti
Northern Europe; Iberia|Klasik/Modern Klasik; Tel/Enstrümantal|Dinginlik|Odak; Gece|Baroque; Early Music; Renaissance Guitar
Balkans|Klasik/Modern Klasik; Vokal/Şarkı|Romantik|Gece|Yunan Entekno; Film Müziği; Politik Şarkı
West Africa|Reggae/Dub|Karanlık|Dans|Afrika Reggae'si; Politik Söz; Sürgün
North America|Caz; Klasik/Modern Klasik|Dinginlik|Odak; Gece; Yolculuk|Béla Fleck; Gershwin/Ferde Grofé Reinterpretation
East & Central Asia|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Yolculuk|Japanese Song; Experimental Ensemble; World Pop
Western Europe|Vokal/Şarkı; Caz|Melankoli|Gece; Yolculuk|Contemporary Folk; Swing; Old-Time Blues; Songwriting
Western Europe|Psikedelik/Rock|Melankoli|Dans; Odak|Minimal Rock; Fransızca Söz; Tekrar Estetigi
Western Europe|Gelenek; Tel/Enstrümantal|Melankoli|Gece|Yidiş/Klezmer; Çello-Keman İkilisi; Balkan Ezgileri
Mediterranean Europe|Reggae/Dub|Neşe|Dans; Yolculuk; Spor|Tarantella Dub; Salento Folk; Elektronik
Western Europe; Levant|Caz|Neşe|Dans; Odak|Caz Dörtlüsü; Ortadoğu Ezgileri; Paris
Levant|Gelenek; Tel/Enstrümantal|Dinginlik|Gece|Buzuk + Filistin Repertuvarı; Al Kamandjâti Çevresi
Caucasus|Gelenek|Neşe|Gece|Kafkas-Anadolu Ortak Repertuvarı; Duduk/Tel
Levant|Tel/Enstrümantal|Neşe|Odak; Spor|Buzuk; Çağdaş Arap Enstrümantalı; Beyrut
Northern Europe|Funk/Soul; Psikedelik/Rock|Karanlık|Yolculuk; Spor|Afrobeat/Funk; Progresif-Psikedelik; Pori Kolektifi
Western Europe; Iberia; Southern Africa; Fusion|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Gece|Art Pop; Electronica; Saudade; Alternative Latin; Afro-Pop
West Africa|Tel/Enstrümantal; Vokal/Şarkı|Dinginlik|Odak; Yolculuk|Kora; Griot Geleneği; Saint-Louis
Mediterranean Europe|Psikedelik/Rock|Neşe|Odak; Yolculuk; Spor|Psych-Folk; Ritmik Tekrar; Milano
Western Europe|Klasik/Modern Klasik; Tel/Enstrümantal|Melankoli|Odak; Gece|Classical Guitar; Contemporary Interpretation; Film Themes
Western Europe|Caz; Funk/Soul|Neşe|Dans; Odak; Yolculuk; Spor|Ethio-Jazz; Jazz-Funk; Cinematic Groove
North America|Funk/Soul; Vokal/Şarkı|Romantik|Dans|R&B/Soul; Detroit; 50'ler-60'lar
North America|Pop/Şarkı-Yazarı; Tel/Enstrümantal|Neşe|Dans|Indie Folk; Newgrass; Cello; Experimental Roots
West Africa|Psikedelik/Rock|Karanlık|Yolculuk; Spor|Tuareg Rock; Çöl Blues; Psikedelik
Iranian World|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Kurdish Maqam; Kamancheh; Tanbur; Tombak
Western Europe; West Africa; Fusion|Funk/Soul|Neşe|Dans; Odak; Yolculuk|Afro-Soul; Gana Highlife + Berlin Cazı
North America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Dans; Gece; Yolculuk|Old-Time; Appalachian Folk; Country Song; Traditional Ballad
Western Europe; West Africa; Fusion|Klasik/Modern Klasik|Melankoli|Dans; Gece|Kora-Gitar Düeti; Akdeniz Folk; Kamara Füzyon
Latin America|Klasik/Modern Klasik; Pop/Şarkı-Yazarı|Dinginlik|Odak; Gece|Film Müziği; Ronroco; Rock Nacional; Bajofondo
West Africa|Gelenek; Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|Tamasheq Desert Blues; Tuareg Guitar; Vocals
North America|Blues; Vokal/Şarkı|Neşe|Dans|Country Blues; Gospel; Acoustic Ensemble
North America|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Bluegrass; Newgrass; Bebop; Ambient Folk; Chamber Ensemble
North America|Vokal/Şarkı; Gelenek|Melankoli|Gece|Old-Time/Country; Batı Virginia; Sokak Müzisyenliğinden
Latin America; Western Europe; Fusion|Gelenek; Funk/Soul|Neşe|Dans|Pachanga; Psychotropical; Colombian Folklore; Funk
North America; Balkans; Fusion|Gelenek; Vokal/Şarkı|Epik|Spor|Yidiş/Balkan/Roman Repertuvarı; Montreal
North America|Blues; Pop/Şarkı-Yazarı|Melankoli|Gece|Folk-Blues; Parmak Stili Gitar; New Orleans Kökeni
North America|Gelenek; Blues|Melankoli|Dans; Spor|Siyah Telli Çalgı Geleneği; Banjo/Keman; Pulitzer
North America|Elektronik|Neşe|Dans; Spor|Video Oyunu Müziği; Enstrümantal
Western Europe; Latin America; Fusion|Psikedelik/Rock; Elektronik|Neşe|Dans; Spor|İsviçre Deneyseli + Kolombiya Tropikali
Northern Europe|Gelenek; Tel/Enstrümantal|Neşe|Odak; Spor|Norwegian Folk Dance; Family Ensemble; Multi-Instrumental
West Africa; Fusion|Gelenek; Caz|Dinginlik|Odak; Gece; Yolculuk|Kora; World Jazz; West African Song
North America|Caz|Dinginlik|Dans; Odak; Gece|Piyano Üçlüsü/Kuartet; Rock Cover'ları; Minneapolis-New York
Western Europe|Gelenek; Caz|Melankoli|Odak; Gece|İngiliz Folk-Rock; Jansch/Renbourn Gitarları; 60'lar
West Africa; Western Europe|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|Malian Song; Contemporary World Pop; Cross-Cultural Collaboration
Latin America; Western Europe; Fusion|Vokal/Şarkı|Dinginlik|Dans; Gece|Kamera Folk; Çok Dilli Vokal Düeti
West Africa|Vokal/Şarkı; Funk/Soul|Neşe|Odak; Yolculuk|Afro-Küba; Son/Pachanga; Analog Africa
Levant|Caz|Melankoli|Odak; Gece; Yolculuk|Contrabas Caz; Akdeniz Cazı; World Jazz
Eastern Europe|Gelenek|Neşe|Dans; Yolculuk|Podlasie Halk Dansları; Cymbały-Keman; Sınır Bölgesi Repertuvarı
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|İngiliz Folk; Queer Politik Söz; Leeds
Western Europe|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|İngiliz Halk Müziği; Siyah Britanya Tarihi
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|British Folk; Double Bass; Vocal Duet
Iberia|Gelenek; Caz|Karanlık|Dans; Odak; Yolculuk|Flamenko; Deneysel Caz; Elektronik/Deneysel Prodüksiyon
Northern Europe|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Sámi Vokali; Kuzey Sámi Dili
Western Europe|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Alan Kayıtları + Şarkı; Doğa Sesleri; Londra
North Africa|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Arabic Classical; Oud; Percussion
Levant; Iranian World; Fusion|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Oud; Persian Percussion; Modal Improvisation
Levant; Iranian World; North America; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece|Free Improvisation; Guitar; Daf/Tombak
Iranian World|Gelenek; Vokal/Şarkı|Melankoli|Dans|Persian Classical Voice; Setar; Tombak; Contemporary Composition
Western Europe; East & Central Asia; Fusion|Elektronik; Pop/Şarkı-Yazarı|Karanlık|Odak; Spor|Japonca Elektro-Pop; Elektronika; House; Synth-Pop
Caucasus; Western Europe; Levant; Fusion|Caz; Tel/Enstrümantal|Melankoli|Gece; Yolculuk|Ermeni Vokal/Beste; Ud; Ney; Piyano; Arap Makamı; Caz
North Africa|Psikedelik/Rock|Epik|Dans; Spor|Gnawa Rock; Psikedelik; Guembri Füzyon
Western Europe|Funk/Soul|Karanlık|Dans; Odak|Balkan Brass; Afrobeat; Enstrümantal Füzyon
Iberia; Latin America|Elektronik; Vokal/Şarkı|Neşe|Dans; Odak; Yolculuk|Brazilian Remix; Portuguese Song; Dancefloor Production
Levant; Northern Europe; Fusion|Gelenek; Caz|Dinginlik|Dans; Odak|Orta Doğu Makamı; Ud; Klarnet/Ney; Irak-Yahudi ve Kürt Mirası; Doğu Cazı
Latin America|Elektronik; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Dijital Cumbia + Vokal; Cumbia; Elektronik Pop; Yerli Kimlik
Iranian World|Gelenek; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Bushehr Folk; Ney-Anban; Bandari Rhythm; Persian Gulf Music
North Africa|Gelenek; Blues|Neşe|Dans; Yolculuk; Spor|Gnawa Gitarı/Guembri; Bechar; Kadın İcracı
North America|Blues; Vokal/Şarkı|Melankoli|Gece|Alabama Blues; Americana; Hırçın Vokal
Western Europe; Levant; Fusion|Reggae/Dub; Tel/Enstrümantal|Neşe|Dans; Odak|Hijaz Maqam; Oud; Trumpet; Dub
Western Europe|Caz; Psikedelik/Rock|Dinginlik|Dans; Odak; Gece|Sörf-Caz; Bariton Saksofon; Amsterdam
East & Central Asia|Elektronik|Melankoli|Gece|Ambient; Techno; Experimental; Downtempo
West Africa|Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Afro-Küba; Senegal Pop; Mbalax Öncesi
Western Europe|Caz|Dinginlik|Dans; Odak; Gece|Çağdaş Caz Dörtlüsü; Bristol; Portishead Ritim Bölümü
Western Europe; West Africa; Fusion|Gelenek; Caz|Neşe|Yolculuk; Spor|Afro-Jazz; Guitar Dialogue; West African Song
Latin America|Pop/Şarkı-Yazarı; Reggae/Dub|Neşe|Dans|Brezilya Popı; Reggae/Baile Funk; Paris
North America|Pop/Şarkı-Yazarı|Melankoli|Gece|Americana; Folk Şarkı-Yazarlığı
North America|Pop/Şarkı-Yazarı; Gelenek|Dinginlik|Gece|Québec Folk; Şiir Bestelemeleri; Montreal
Western Europe; Latin America; Fusion|Tel/Enstrümantal|Dinginlik|Dans; Odak|Enstrümantal Çöl Gitarı; İki Kardeş; Easy Eye Sound
North America|Gelenek; Vokal/Şarkı|Melankoli|Gece|Appalachian Banjo; Contemporary Folk; Roots Song
North America|Gelenek; Elektronik|Neşe|Dans; Spor|Québec Elektro-Trad; Keman + Ayak Vuruşu
Latin America|Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|Çağdaş MPB; Rio de Janeiro
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Gece|İngiliz Folk; Melankolik Söz Yazarlığı
Western Europe|Gelenek; Vokal/Şarkı|Dinginlik|Odak|Manx Galcesi; Manx Halk Müziği; Soul/Groove Etkileri
Mediterranean Europe|Reggae/Dub; Nefesli/Bando|Melankoli|Dans; Yolculuk; Spor|Dub; Rocksteady; Reggae; Ska-Jazz
Mediterranean Europe|Tel/Enstrümantal|Melankoli|Odak; Gece; Yolculuk|Gitar; Sinematik Enstrümantal; Capossela Çevresi
Mediterranean Europe|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Gece|Neapolitan Song; Mediterranean; Acoustic Trio
Western Europe|Psikedelik/Rock|Karanlık|Dans; Yolculuk; Spor|İngiliz Folk-Rock; Gitar; Şarkı Yazımı
East & Central Asia|Gelenek; Vokal/Şarkı|Dinginlik|Gece|Siberian Folk; Old Russian Song; Historical Instruments; Throat Singing
Caucasus|Tel/Enstrümantal; Gelenek|Neşe|Odak; Spor|Azeri Elektrikli Gitar; Mugam Ezgileri; Bakü
East Africa|Caz|Melankoli|Yolculuk|Etiyopya Cazı; Saksofon; Addis Ababa
East & Central Asia; Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Tibet Repertuvarı; Himalaya Ezgileri; Fransa Yerleşik
Latin America|Gelenek|Romantik|Odak|Tango Orkestrası; Dans Salonu Repertuvarı; Buenos Aires
Mediterranean Europe; Anatolia; Fusion|Psikedelik/Rock|Karanlık|Dans; Yolculuk; Spor|Doğu Akdeniz Psikedelisi; Lefkoşa'nın Bölünmüş İki Yakası; Baş Klarnet + Tzouras
North Africa|Elektronik; Gelenek|Neşe|Dans|Tuareg/Sufi İlahileri; Trap; Raï; Güney Libya Ritimleri
Southern Africa; West Africa; Western Europe; Fusion|Gelenek; Klasik/Modern Klasik|Neşe|Dans; Yolculuk; Spor|Çello + Batı Afrika Perküsyonu + Yaylı Kadrosu; Abel Selaocoe Projesi; Cello; West African Percussion; Contemporary Chamber
Central Africa|Elektronik|Neşe|Dans; Spor|Hurda Enstrüman; DIY Elektronik; Kinşasa
North America|Gelenek; Funk/Soul|Melankoli|Spor|Mississippi Gospel-Funk; Aile Topluluğu
Western Europe|Caz; Elektronik|Dinginlik|Dans; Odak; Yolculuk|Hang/Handpan; Ambient Caz; Londra
North America|Caz; Funk/Soul|Neşe|Odak|Piyano-Vokal; Sivil Haklar; Klasik Eğitim
South Asia|Klasik/Modern Klasik; Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Sarod; Hindustani Klasik; Maihar Ekolü
Levant|Caz|Dinginlik|Odak; Yolculuk|Contrabas Caz; Akdeniz Cazı; World Jazz
Western Europe|Caz|Neşe|Odak|Jazz Manouche; Django Geleneği
Western Europe|Caz|Romantik|Dans; Odak|Gypsy Jazz; Manouche Gitar; Django Geleneği
Western Europe|Psikedelik/Rock|Dinginlik|Dans; Odak|Cumbia-Surf; Enstrümantal Psikedelik; Anadolu Etkisi
Northern Europe; Eastern Europe; Fusion|Caz; Gelenek|Neşe|Spor|Klezmer/Balkan + Nordik Caz; Kontrbas Liderliği
Western Europe|Gelenek; Vokal/Şarkı|Romantik|Dans; Gece|İngiliz Folk Üçlüsü; Gitar-Akordeon-Keman
Mediterranean Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece|Diyatonik Akordeon + Klarnet; Toskana Hattı
Iranian World; North America; Fusion|Gelenek; Caz|Dinginlik|Yolculuk|Fars Klasik Vokali + Flüt; Golestan Topluluğu
Southeast Asia|Gelenek; Reggae/Dub|Neşe|Dans; Odak; Spor|Molam; Phin; Khaen; Dub; Thai Funk
South Asia|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Urduca Şiir + Minimal Düzenleme; Gazel Esini
Oceania|Funk/Soul; Vokal/Şarkı|Neşe|Yolculuk|Indigenous Australian Soul; Blues; Analog Rhythm Section
Levant|Elektronik; Caz|Karanlık|Odak; Yolculuk; Spor|Arap Makamı; Acid Jazz; Progresif Rock; Canlı Elektronik
Western Europe; Levant; Fusion|Caz; Vokal/Şarkı|Melankoli|Odak|Trombon + Arapça Vokal; Çağdaş Caz
Levant; East Africa; Fusion|Caz; Funk/Soul|Neşe|Dans; Odak; Spor|Ethio-Jazz; Afrobeat; Psychedelic Mediterranean Funk; Brass; Tezeta Modes
Anatolia|Psikedelik/Rock; Funk/Soul|Neşe|Odak|Anatolian Psychedelia; Space Funk; Alternative Rock; Electronic Arrangement
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Dinginlik|Odak; Gece; Yolculuk|Meksika Halk Şarkısı; Son Jarocho; Bolero
South Asia|Vokal/Şarkı; Gelenek|Epik|Yolculuk; Spor|Kavvali; Nusrat Geleneği; Sufi Repertuvar
Levant|Gelenek|Karanlık|Yolculuk|Deneysel Arap Müziği; Serbest Doğaçlama; Noise-Folk
Balkans|Nefesli/Bando|Neşe|Dans; Odak; Spor|Balkan Brass; Roman Nefesli Orkestrası
Western Europe; Iranian World; Balkans; Levant; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Zarb + Çello + Piyano; Çok Gelenekli Oda Müziği
Western Europe|Hip-Hop/Rap|Karanlık|Dans; Spor|İrlandaca Rap; Belfast; Politik Söz
Western Europe; Levant; Anatolia; Fusion|Klasik/Modern Klasik; Gelenek|Neşe|Odak|Ud/Keman/Kontrbas/Perküsyon; Makam + Avrupa Klasiği; Berlin
Southern Africa|Funk/Soul; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk; Spor|Chigiyo; Zimbabwean Funk; Reggae; Sungura; Dance Band
Balkans|Elektronik; Gelenek|Neşe|Dans; Yolculuk; Spor|Balkan Elektro; Kaval/Tel; Bükreş
Balkans|Caz|Romantik|Gece|Sevdalinka; Bosna Halk Müziği; Caz Düzenleme
Northern Europe; West Africa; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Dans|Scottish Gaelic Song; Kora; Contemporary Folk
Iberia|Klasik/Modern Klasik|Melankoli|Gece|Polifonik Vokal Düeti; Akdeniz İlahileri; Erken Müzik; Catalonia
Western Europe; South Asia; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans|Punjabi Folk; Bhangra; Tabla; Dholak; Contemporary Production
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Gece|Latin Alternatif; Bolero/Ranchera Etkisi; Meksika Yerleşik
Western Europe|Psikedelik/Rock|Neşe|Dans; Odak|Cumbia-Surf; Enstrümantal Psikedelik; Anadolu Etkisi
North Africa|Vokal/Şarkı|Neşe|Dans|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
South Asia; Oceania; Fusion|Gelenek; Caz|Neşe|Odak|Baul; Indo-Jazz; Bengali Folk; Improvisation
Levant; Anatolia; Fusion|Gelenek; Klasik/Modern Klasik|Dinginlik|Dans; Odak; Gece|Ud; Arp; Perküsyon; Osmanlı/Sefarad ve Çağdaş Oda Müziği
Levant|Tel/Enstrümantal; Pop/Şarkı-Yazarı|Melankoli|Gece|Instrumental Guitar; Bouzouki; Mediterranean Ensemble; Live Looping
East & Central Asia|Gelenek; Psikedelik/Rock|Neşe|Yolculuk|Amis Indigenous Song; Psychedelic Rock; Experimental Funk; Improvisation
West Africa; Mediterranean Europe; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Batı Afrika Ritimleri; Wolof/Senegal Etkisi; Kültürlerarası Füzyon
North America|Funk/Soul; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Afrobeat; Brooklyn; Fela Geleneği
Eastern Europe|Gelenek; Tel/Enstrümantal|Melankoli|Dans; Gece|Tango Nuevo; Bandoneón Kuarteti
West Africa|Gelenek; Vokal/Şarkı|Melankoli|Gece|Senegalese Xalam; Wolof Song; Afrobeat
Western Europe; North America; Fusion|Gelenek; Caz|Neşe|Dans; Odak|Folk Traditions; Bluegrass; Acoustic Dialogue
Western Europe; Caribbean|Reggae/Dub|Neşe|Yolculuk|Lovers Rock; İngiliz Reggae'si; Londra
North America; Caribbean; West Africa; Fusion|Caz; Funk/Soul|Neşe|Yolculuk; Spor|Afro-Jazz; Haitian Roots; Ghanaian Percussion; Improvisation
North America|Gelenek; Vokal/Şarkı|Neşe|Yolculuk|Klezmer; Yiddish Song; Jewish Roots
Levant; Iberia; Fusion|Gelenek; Klasik/Modern Klasik|Melankoli|Gece; Yolculuk|Syriac Hymn; Cantigas de Santa Maria; Arabic Maqam; Medieval Iberian Music
Southeast Asia|Psikedelik/Rock|Neşe|Dans; Odak; Spor|Endonezya Psychedelia; Dangdut-Funk; Enstrümantal
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Çağdaş Caz Dörtlüsü; Davul; Paris-New York
Caribbean; North America; Fusion|Funk/Soul; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Cuban Conga; Latin Dance-Pop; Miami Sound
Latin America|Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Latin Pop; Worldbeat
North America|Gelenek; Vokal/Şarkı|Melankoli|Dans; Gece; Yolculuk|Opera Sopranosu + Banjo + Country; Appalachian Köprüsü
Oceania|Funk/Soul; Vokal/Şarkı|Melankoli|Yolculuk|Aborjin Soul; Gumbaynggirr Kökeni
Mediterranean Europe|Caz; Nefesli/Bando|Neşe|Dans|Napoli Geleneği + Caz; Politik Şarkı; Napoli
North Africa|Psikedelik/Rock|Neşe|Odak; Spor|Tuareg Rock; Çöl Blues; Assouf
West Africa; Western Europe|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk|Yoruba Perküsyonu; Afro-Funk; Londra
Southern Africa|Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Parmak Stili Gitar; Enstrümantal; Durban
Southern Africa; North America; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Akustik Gitar İkilisi; Parmak Stili
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Contemporary Jazz; Free Improvisation; Bass/Clarinet/Drums
Western Europe|Caz; Elektronik|Karanlık|Odak|Placebo (BE); Telex; Jazz-Funk/Synth
Western Europe|Elektronik|Karanlık|Odak|Fransız Elektroniği; Keman + Sample
Western Europe|Tel/Enstrümantal|Dinginlik|Odak; Gece|Solo Elektrikli Gitar; Enstrümantal; Yorkshire
North America|Pop/Şarkı-Yazarı|Dinginlik|Odak; Gece|Ambient Folk; Teksas; Alan Kayıtları
West Africa|Gelenek; Klasik/Modern Klasik|Melankoli|Yolculuk|Kora; West African Vocal Traditions
Caribbean|Caz|Neşe|Dans; Spor|Çağdaş Küba Cazı; Bebop; Afro-Küba Öğeleri
North America; Caribbean|Blues; Pop/Şarkı-Yazarı|Melankoli|Gece|New Orleans Folk-Blues; Haiti Kökeni
Latin America|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Yolculuk|Afro-Peruvian; Landó; Festejo
North America|Vokal/Şarkı; Gelenek|Melankoli|Dans; Gece; Yolculuk|Child Ballad'ları; Amerikan Folk Düeti
Western Europe; Latin America; Fusion|Caz; Vokal/Şarkı|Dinginlik|Odak; Gece; Yolculuk|World Jazz; Clarinet; Piano; Cello; Beatbox; Latin American and Balkan Influences
Eastern Europe|Vokal/Şarkı|Neşe|Gece; Yolculuk|Roman Repertuvarının Slovak Yorumcusu; 'Manuşa' (2022)
Iberia; Latin America; Fusion|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Gece|Latin Pop; Música Mexicana; Sierreño
Caucasus; Anatolia; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Armenian Duduk; Turkish Saz; Folk Dialogue
Iranian World|Tel/Enstrümantal|Romantik|Dans; Yolculuk; Spor|Kamançe/Keman; Fars Klasik; Berlin
East & Central Asia|Psikedelik/Rock; Gelenek|Romantik|Gece|Uygur Folk-Rock; Sovyet Estrada; 1978 Kaydı; Muqam Mirası
North America; Levant; Fusion|Caz; Gelenek|Neşe|Odak|Irak Makamı + Caz; Trompet/Santur; Çeyrek Ses
Latin America|Pop/Şarkı-Yazarı|Dinginlik|Gece|Dream Pop; Latin Renkler; Londra
Western Europe|Pop/Şarkı-Yazarı|Dinginlik|Odak; Gece|İrlanda Folk + Elektronik; Belfast
North Africa|Caz|Dinginlik|Odak; Gece; Yolculuk|Oud; Arap Klasik Müziği; Kamara Caz
Caribbean|Reggae/Dub; Caz|Neşe|Dans; Odak; Yolculuk|Ska/Reggae Gitarı; Stüdyo Sesyonları; Studio One
Western Europe; South Asia; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Punjabi Folk; Bhangra; Raag Bhairavi; Contemporary Production
Anatolia|Elektronik; Tel/Enstrümantal|Dinginlik|Odak; Gece|Anadolu Enstrümantal; Elektronik Folk
North America|Caz; Vokal/Şarkı|Melankoli|Gece|Caz Vokali; New Orleans
North Africa|Hip-Hop/Rap; Elektronik|Neşe|Dans|Arabic Hip-Hop; Egyptian Hip-Hop; Mahraganat
Levant|Tel/Enstrümantal; Gelenek|Melankoli|Dans; Odak; Gece|Ud; Irak Halk Ezgileri; Enstrümantal Maqam
Western Europe; Fusion|Klasik/Modern Klasik; Caz|Dinginlik|Odak; Gece|Cello/Piano Duo; Chamber Improvisation; Contemporary Classical
Latin America|Gelenek; Nefesli/Bando|Dinginlik|Odak; Gece|Quena/And Flütleri; Tango-Milonga Hattı
West Africa|Blues; Psikedelik/Rock|Neşe|Odak; Yolculuk; Spor|Tuareg Çöl Blues'u; Kidal; Tamashek Söz
Western Europe|Elektronik|Neşe|Odak; Spor|Downtempo; Breakbeat; Nu Jazz; Sample-Based Electronica
Levant|Caz|Neşe|Dans; Odak; Spor|Free Jazz; Gnawa; Spiritual Jazz; Gimbri
Northern Europe; Eastern Europe; North Africa; Fusion|Caz; Gelenek|Melankoli|Gece; Yolculuk|Microtonal Folk-Jazz; Arabic Maqam; Norwegian Folk
North America; East Africa; Levant; Fusion|Caz; Funk/Soul|Neşe|Dans; Odak|Ethiopian Jazz; Afrobeat; Free Jazz; Middle Eastern Modes; Minimal Groove
Western Europe; Latin America; Fusion|Elektronik; Psikedelik/Rock|Neşe|Dans|Indie; Dub; Experimental Pop
Latin America|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Ranchera; Bolero; Çıplak Gitar-Vokal
Eastern Europe|Gelenek|Neşe|Dans; Spor|Podlasie Halk Dansları; Cymbały-Keman; Sınır Bölgesi Repertuvarı
Levant; Western Europe|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Odak|Arabic Electronica; Trip-Hop; Alternative Pop; Multidialectal Arabic
Levant|Funk/Soul|Neşe|Dans; Spor|Beyrut Library Funk; Enstrümantal; Arap Groove
Arabian Peninsula|Psikedelik/Rock|Karanlık|Yolculuk; Spor|Suudi Psikedelik Rock; Kadın Grubu; Arap Retro
West Africa|Blues; Psikedelik/Rock|Epik|Dans; Spor|Tuareg Çöl Blues'unun Kurucuları; Tamashek; Assouf
Anatolia|Psikedelik/Rock|Karanlık|Yolculuk|Anadolu Psych; Sürrealist Sahne Kimliği; İstanbul
North America|Gelenek; Tel/Enstrümantal|Melankoli|Dans; Gece|Bluegrass; Kadın Tel Topluluğu
North America; Mediterranean Europe; Fusion|Gelenek|Epik|Gece; Yolculuk|Amerikan Kök Müziği; Banjo; Akdeniz Çerçeve Davulu
Western Europe|Caz; Elektronik|Dinginlik|Odak|Gal Halk Ezgileri; Dub; Kontrbas
Anatolia|Elektronik|Romantik|Dans; Odak; Spor|Anadolu Psychedelia; Funk; Elektronik Folk
Levant|Elektronik; Caz|Neşe|Odak; Yolculuk; Spor|Arap Şaabi; Free Jazz; Elektroakustik
Iranian World|Elektronik|Neşe|Dans; Odak; Spor|IDM; Downtempo; Ambient; Drum and Bass; Persian Modal Melodies; Dub
Anatolia|Psikedelik/Rock; Vokal/Şarkı|Neşe|Dans; Odak|Anadolu Psikedelyası; Çağdaş İstanbul Sahnesi
Levant|Tel/Enstrümantal|Dinginlik|Dans; Odak|Oud; Arap Makamı; Çağdaş Beste
Western Europe; Eastern Europe; Fusion|Elektronik; Hip-Hop/Rap|Neşe|Dans; Spor|Parodik Sovyet Estetiği; Balkan/Slav Sample'ları; Paris
Caucasus; Western Europe; Fusion|Gelenek|Dinginlik|Dans|Ermeni Halk Müziği; Balkan Füzyon; Pop
Southeast Asia|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece; Yolculuk|Endonezya Etnik Füzyon; Çello; Bali/Ubud
Iberia|Vokal/Şarkı|Melankoli|Gece; Yolculuk|Fado; Portekiz Vokal Geleneği
Mediterranean Europe|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Spor|Napoli Lehcesi; Elektronik + Halk Ezgisi
North America; Iranian World; West Africa; Fusion|Klasik/Modern Klasik|Dinginlik|Gece; Yolculuk|Fars-Ortaçağ Füzyonu; Kora; Erken Müzik
Central Africa; West Africa; Caribbean; Fusion|Elektronik; Reggae/Dub|Neşe|Dans; Spor|Afrofuturist Dub; Mbalax; Sahelian Electronic
Caribbean; Iberia; Fusion|Funk/Soul; Gelenek|Neşe|Dans|Afro-Küba; Latin Soul/Funk; Son; Timba; Boogaloo
Western Europe; West Africa; Fusion|Funk/Soul; Gelenek|Melankoli|Dans|Afro-Funk; Desert Blues; Cross-Continental Collaboration
Levant; Anatolia; Caucasus; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Mediterranean Instrumental; Oud; Kanun; Duduk; World Fusion
Anatolia; Northern Europe; North America|Caz; Funk/Soul|Karanlık|Dans; Odak; Spor|Istanbul Jazz-Funk; Saxophone; Percussion
Anatolia|Elektronik; Caz|Karanlık|Odak; Yolculuk|Caz; Elektronik; Nublu Sahnesi
Western Europe|Caz; Elektronik|Dinginlik|Odak; Yolculuk|Flüt + Efekt; Çağdaş Fransız Cazı
Western Europe|Caz; Elektronik|Dinginlik|Odak; Gece; Yolculuk|Flüt + Efekt; Çağdaş Fransız Cazı
North Africa|Vokal/Şarkı|Neşe|Dans; Spor|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
Eastern Europe; Western Europe|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Keman/Besteci; Ukrayna Halk Ezgileri; Londra
Levant|Vokal/Şarkı; Gelenek|Neşe|Dans; Odak; Spor|Arap Usulleri (malfuf); İbrani/Arap Vokal
Western Europe|Hip-Hop/Rap|Neşe|Dans; Odak; Spor|Trip-Hop; Sample Bazlı Enstrümantal Hip-Hop
Western Europe|Caz|Melankoli|Yolculuk|Minimal Caz; Fısıltılı Vokal; Flüt
Levant; Arabian Peninsula; Fusion|Elektronik; Gelenek|Epik|Dans; Spor|Irak Geleneksel Şarkısı; Yemen-Yahudi Mirası; Analog Synth; Elektronik Groove
Western Europe|Psikedelik/Rock; Pop/Şarkı-Yazarı|Melankoli|Odak|İsviçre Almancası Vokal; Lo-Fi Psych
West Africa|Gelenek; Vokal/Şarkı|Neşe|Dans|Morna; Yeşil Burun Geleneksel Şarkısı
Oceania; West Africa; Fusion|Funk/Soul; Caz|Epik|Odak; Spor|Cinematic Soul; Afrobeat; Psychedelia; Nigerian Heritage; Analog Production
Latin America; Western Europe; Fusion|Caz; Klasik/Modern Klasik|Neşe|Dans; Odak; Yolculuk|Brazilian Jazz; Orchestral Jazz Fusion
Western Europe|Vokal/Şarkı; Caz|Melankoli|Gece|Chamber Folk; Jazz Ballad; Art Song
Northern Europe|Pop/Şarkı-Yazarı|Melankoli|Yolculuk|İskandinav Şarkı-Yazarı; Orkestral Düzenleme
North America|Blues; Vokal/Şarkı|Karanlık|Dans; Yolculuk|Delta Bluesu; Film Müziği Kaydı
Latin America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Gece|Kolombiya Halk Şarkısı; Toplumsal Söz; Tiple/Gitar
Iberia; Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Indie Folk; Şarkıcı-Şarkı Yazarı; Klasik Müzik Altyapısı
Levant; Western Europe; Fusion|Psikedelik/Rock; Caz|Karanlık|Spor|Arap Vokal + Elektrikli Grup; Lyon
Northern Europe; Caucasus; Western Europe|Caz; Tel/Enstrümantal|Dinginlik|Odak; Gece|Contemporary European Jazz; Bass; Piano; Guitar; Improvisation
Northern Europe|Caz|Dinginlik|Odak; Yolculuk|Saksofon; ECM Estetigi; Nordik Caz
Western Europe|Elektronik; Caz|Karanlık|Odak|Sinematik Elektronik; Alan Kayıtları; Edinburgh
North America|Funk/Soul|Melankoli|Dans; Odak; Spor|Latin Soul/Funk; Black Pumas; Austin
North America|Funk/Soul|Epik|Spor|Psikedelik Soul; Austin; Adrian Quesada Prodüksiyonu
North America; Western Europe|Caz; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Caz Vokali; Arkaik Repertuvar Araştırması; Miami/Paris
Western Europe; Northern Europe; Fusion|Caz|Neşe|Odak|Piyano Üçlüsü; İskandinav + Kuba Ritmi; Hamburg
Oceania|Funk/Soul; Caz|Neşe|Dans; Odak|Enstrümantal Sinematik Soul; Boogaloo; 'Pub Soul'; Yedili + Nefesliler
Oceania|Caz; Funk/Soul|Epik|Dans; Yolculuk; Spor|Afrobeat; Spiritual Jazz; Geniş Kadro; Adelaide
Southern Africa|Gelenek; Funk/Soul|Epik|Dans; Spor|Soweto; Ritmüel Vokal/Perküsyon; Uzun Formlu İcra
Mediterranean Europe|Gelenek|Neşe|Yolculuk|Pizzica; Tarantella; Salento Halk Müziği
Latin America; Western Europe; Fusion|Gelenek; Elektronik|Neşe|Dans; Odak|Afro-Kolombiya Pasifik Geleneği; Marimba; Perküsyon; Elektronik/Güçlendirilmiş Füzyon
North America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Gece|Old-Time; Appalachia; Clawhammer Banjo-Fiddle
North America; Iranian World; Fusion|Elektronik|Neşe|Dans|Psikedelik Beat; Sample Prodüksiyon; Fars Etkisi
Northern Europe; Anatolia; Fusion|Psikedelik/Rock; Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Krautrock/Trâns; Saz + Darbuka + Ağız Kopuzu; Techno/Acid; Göteborg
Northern Europe|Pop/Şarkı-Yazarı|Melankoli|Odak; Gece|İskandinav Folk; Parmak Ucı Gitar; Sade Prodüksiyon
Mediterranean Europe|Tel/Enstrümantal; Gelenek|Melankoli|Odak; Gece|Sardinya Telleri; Chitarra Battente / Mandola
Mediterranean Europe|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Sicilya Lehçesi; Cantu e Cuntu; Akdeniz Kök Müziği; Rosa Balistreri Hattı
East & Central Asia; North America; Fusion|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans; Odak; Spor|Indie Pop; Sample-Based Production; Global Groove
Iranian World|Vokal/Şarkı; Gelenek|Romantik|Dans; Yolculuk; Spor|İran Vokali; Tehrani Drom Topluluğu
Western Europe; East Africa; Fusion|Reggae/Dub|Neşe|Dans; Yolculuk; Spor|Etiyopya Ezgileri + Dub Prodüksiyonu; Addis
North America|Psikedelik/Rock|Neşe|Spor|Gypsy Punk; Balkan Rock; Göçmen Kabare
Iranian World; South Asia; North America; Fusion|Tel/Enstrümantal; Caz|Dinginlik|Odak; Yolculuk|Santur; Harp; Tabla; Chamber Jazz
Western Europe|Elektronik; Klasik/Modern Klasik|Epik|Odak|Piyano + Davul İkilisi; Post-Rock/Elektronik
Western Europe|Caz|Karanlık|Dans; Odak|Fransız Piyano Üçlüsü; Çağdaş Caz
Western Europe|Caz; Elektronik|Neşe|Odak; Spor|Londra Caz Sahnesi; Davul; Grime/Afrobeat Etkisi
West Africa|Funk/Soul; Gelenek|Epik|Dans; Odak; Yolculuk; Spor|Highlife + Afrobeat; Gitar; Accra
West Africa|Tel/Enstrümantal; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Kora; Griot Geleneği; Kelt İşbirlikleri
West Africa; Iberia; Fusion|Tel/Enstrümantal|Neşe|Odak; Spor|Kora-Çello Düeti; Griot Geleneği; İber Vokal
Latin America|Caz|Romantik|Odak; Gece|Jazz Manouche; Django Geleneği; Los Angeles
East & Central Asia|Caz; Funk/Soul|Karanlık|Dans; Yolculuk; Spor|Japon Rare Groove; Caz-Funk; Soul; Sanshin; Japon Geleneksel Şarkısı
North America|Psikedelik/Rock|Karanlık|Odak; Yolculuk|Math Rock; Mikrotonal Gitar; Enstrümantal Prog; Quebec
Latin America; North America; Fusion|Gelenek|Neşe|Dans; Odak; Spor|Müzik Llanera + Bluegrass; Arpa/Banjo İkilisi
Latin America|Gelenek; Funk/Soul|Neşe|Dans; Spor|Tropikal/Cumbia; Bogotá Sahnesi
Iberia|Gelenek|Romantik|Yolculuk|Ball Folk / Ball de Plaça; Flüt-Keman-Diyatonik Akordeon; Katalonya; Catalonia
Western Europe|Gelenek|Neşe|Dans; Spor|Poitou Balfolk; Akordeon/Keman; Dans Müziği
Balkans|Caz|Melankoli|Yolculuk; Spor|Sevdah; Roman Müziği; Balkan Caz
East Africa|Caz|Neşe|Odak; Yolculuk; Spor|Ethio-Jazz'ın Kurucusu; Vibrafon; Addis Ababa
Northern Europe; West Africa; Fusion|Tel/Enstrümantal; Gelenek|Neşe|Odak|İsveç Flütü + Ngoni; Mande Geleneği; Sınırlararası Akustik Düet
Levant; North America; Fusion|Elektronik; Gelenek|Melankoli|Dans; Yolculuk|Middle Eastern Folk; Electronic Remix; Cross-Cultural Groove
North Africa; North America|Gelenek; Vokal/Şarkı|Neşe|Dans; Gece|Chaâbi + Flamenko/Rumba; Montreal
Western Europe; North America; Fusion|Caz; Elektronik|Dinginlik|Odak; Gece|Piano Fusion; Modern Jazz; Ambient Groove
Northern Europe; Levant; Fusion|Vokal/Şarkı; Caz|Dinginlik|Yolculuk|Muvaşşah; Oud-Vokal; Mikrotonal Piyano
Western Europe; Anatolia; Iranian World; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Yolculuk|Çello + Bağlama + Perküsyon; Making Tracks Rezidansı
West Africa|Blues; Psikedelik/Rock|Neşe|Dans; Yolculuk; Spor|Çöl Blues; Timbuktu'dan Sürgün; Bamako
North America|Blues; Gelenek|Melankoli|Gece; Yolculuk|Piedmont Blues/Gospel; Parmak Stili Gitar; Harlem
Western Europe|Caz|Dinginlik|Odak; Gece; Yolculuk|Spiritual Jazz; Saksofon + Şiir; Manchester/Londra
Southeast Asia|Gelenek; Tel/Enstrümantal|Neşe|Spor|Endonezya Etnik Füzyon; Çoklu Enstrüman; Ses İyileştirme; Ubud/Bali
North America; Latin America; Fusion|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans|Art Pop; Latin Alternative; Experimental
Western Europe|Klasik/Modern Klasik; Gelenek|Melankoli|Odak; Gece|Fransız Orkestral Aranjman; Akordeon + Mandolin
Western Europe|Caz|Epik|Odak; Yolculuk|Britanya Cazı; Karayip Ritmi; Politik Enstrümantal
Western Europe|Caz; Elektronik|Melankoli|Dans; Odak; Gece|İskandinav Cazı + Turntablism; Piyano-Theremin; Bristol Sahnesi
Iberia; Levant; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak; Gece; Yolculuk|Arabic-Andalusian; Ney; Oud; Qanun
North America|Klasik/Modern Klasik; Tel/Enstrümantal|Dinginlik|Odak; Gece|Klasik Gitar; Besteci-İcracı; LAGQ
Western Europe|Tel/Enstrümantal|Epik|Dans; Odak; Yolculuk; Spor|Kaval; Balkan-Anadolu Nefesli Geleneği
Iranian World|Vokal/Şarkı; Gelenek|Neşe|Dans|Fars Halk Repertuvarı; Çağdaş Düzenleme; Tahran
Western Europe|Psikedelik/Rock|Karanlık|Dans; Odak; Spor|Krautrock; Deneysel Rock; Döngüsel Ritim
Western Europe; North America; Fusion|Elektronik; Pop/Şarkı-Yazarı|Neşe|Dans|Art Pop; Ambient; Experimental
North America|Gelenek; Blues|Karanlık|Dans|Siyah Appalachia Geleneği; Banjo/Keman; Arşiv Çalışması
North America|Gelenek; Blues|Karanlık|Dans|Siyah Appalachia Geleneği; Banjo/Keman; Arşiv Çalışması
North America|Gelenek; Vokal/Şarkı|Neşe|Dans|Bluegrass; Flatpicking; Americana; Jamgrass; Şarkıcı-Şarkı Yazarı
North Africa; West Africa|Pop/Şarkı-Yazarı; Gelenek|Neşe|Yolculuk|Moritanya Pop'u; 1984–1987 Kayıtları; Arap-Berberi Gelenek
Latin America|Reggae/Dub; Gelenek|Neşe|Dans|Reggae-Cumbia; Mayan Roots; Socially Conscious Lyrics; Pre-Hispanic Instruments
Latin America|Elektronik|Neşe|Dans|Meksika Elektronik/Cumbia; Sample Bazlı; Camilo Lara
Western Europe|Vokal/Şarkı; Gelenek|Melankoli|Gece|Oksitan Kadın Vokal Üçlüsü; Tef; Toulouse
Iberia|Gelenek; Tel/Enstrümantal|Dinginlik|Dans; Odak; Yolculuk|Flamenko Arpı; İlk Kadın Flamenko Arpisti
Iberia|Vokal/Şarkı; Gelenek|Dinginlik|Dans; Yolculuk|Katalan/Balear Halk Şarkısı; Deneysel Vokal
Caribbean; Latin America; Western Europe; Fusion|Hip-Hop/Rap; Funk/Soul|Neşe|Yolculuk|Latin Hip-Hop; Brazilian Soul; Afrobeat; Reggae; Bilingual Vocals
North America|Pop/Şarkı-Yazarı|Dinginlik|Gece|Güney Amerikana; Indie Rock; Kentucky Şarkı Yazımı
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Flemish Folk; Arabic-Inspired Song; Acoustic Trio
Levant|Caz|Neşe|Dans; Odak|Çeyrek Tonlu Trompet; Arap Makamı + Caz
Latin America|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Dans; Odak|Kolombiya Halk Şarkısı; Toplumsal Söz; Tiple/Gitar
Levant; Western Europe; South Asia; Fusion|Psikedelik/Rock|Dinginlik|Odak; Spor|Kavvali; Racastan Halk Müziği; Sufi-Rock Füzyon
Western Europe; Latin America; Fusion|Elektronik; Reggae/Dub|Neşe|Dans; Spor|Brazilian Pop; Dub; Global Bass; Portuguese Vocal
North Africa|Elektronik|Melankoli|Gece|Arap Alternatif; Protest Şarkı; Elektronik
Latin America|Klasik/Modern Klasik; Caz|Dinginlik|Odak|Gitar/Piyano; ECM; Amazon Esini
Western Europe|Caz; Tel/Enstrümantal|Neşe|Dans; Odak; Gece|Gypsy Jazz; Manouche Guitar; Acoustic Trio
Western Europe|Caz; Tel/Enstrümantal|Neşe|Dans; Odak|Gypsy Jazz; Jazz Manouche; Guitar Trio
Levant|Psikedelik/Rock|Neşe|Odak; Spor|Psychedelic Rock; Akdeniz Surf; Doğu Rock
Levant|Psikedelik/Rock; Funk/Soul|Neşe|Dans; Odak|Anadolu Psych; Saz; Dabke; Bandari; Dub; Batov Records
East & Central Asia|Gelenek; Vokal/Şarkı|Melankoli|Odak|Altai Mongolian Folk; Tuvan Music; Throat Singing; Morin Khuur; Tovshuur
Western Europe|Caz; Tel/Enstrümantal|Karanlık|Dans; Odak|Arp Merkezli Caz; Bastard Jazz Çevresi
Northern Europe|Caz; Elektronik|Dinginlik|Dans; Odak|İsveç Cazı/Sinematik; Björn Yttling
Fusion|Elektronik|Neşe|Dans; Odak|Deneysel Elektronik; Enstrümantal
Latin America|Hip-Hop/Rap; Funk/Soul|Neşe|Yolculuk|Brezilya Rap'i; Planet Hemp Çevresi; Rio
North America; Latin America; Western Europe; Fusion|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans|Tropical Electronic; Latin Alternative
North America; Latin America; Fusion|Caz; Gelenek|Neşe|Odak|Choro; Latin Jazz; Jazz-Funk
Latin America|Gelenek; Elektronik|Neşe|Dans; Odak|Gaita + Trompet; Karayip Kolombiya Hattı
Western Europe|Elektronik|Karanlık|Odak; Yolculuk|Trip-Hop; Bristol Sound; Sample Bazlı Prodüksiyon
East & Central Asia|Pop/Şarkı-Yazarı; Gelenek|Romantik|Yolculuk|Hailufeng Lehçesi; Guangdong; Halk-Rock Anlatısı
East & Central Asia|Vokal/Şarkı; Gelenek|Melankoli|Dans|Chinese Folk Song; Indie Folk; Ensemble Vocals
Eastern Europe|Gelenek; Caz|Dinginlik|Odak; Gece|Hungaro-Serbian Tambura; Balkan Folk; Jazz Fusion
Mediterranean Europe; North America|Elektronik; Caz|Dinginlik|Dans; Odak; Gece|Kolektif Elektronik + Saksofon/Nefesli
Balkans; North America; Fusion|Psikedelik/Rock; Nefesli/Bando|Karanlık|Dans|Balkan Punk/Metal; Seattle; Göç Anlatısı
East & Central Asia|Psikedelik/Rock|Dinginlik|Dans; Odak|Japon Psych; Kikagaku Moyo Çevresi; Guruguru Brain
West Africa|Blues; Funk/Soul|Neşe|Dans; Odak; Yolculuk; Spor|Desert Blues; Guitar; West African Groove
Western Europe|Caz; Tel/Enstrümantal|Dinginlik|Odak; Gece|Chamber Jazz; Double Bass; Saxophone; Minimalism; Cinematic Folk
Western Europe; North America; Fusion|Caz; Gelenek|Dinginlik|Dans; Odak|Gypsy Jazz; Guitar Duo; Parisian Swing
Anatolia|Psikedelik/Rock|Neşe|Spor|Anadolu Psych; Bağlama/Saz; Hamburg
Western Europe|Pop/Şarkı-Yazarı; Elektronik|Dinginlik|Dans; Odak|Piyano-Chanson; Nôze Çevresi; Paris
Central Africa|Vokal/Şarkı; Funk/Soul|Neşe|Gece|Congolese Rumba; Soukous Roots; Guitar; Lingala Song
West Africa|Blues; Funk/Soul|Neşe|Dans; Spor|Afrobeat; Highlife; Blues
Balkans|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece|Yunan Lavtası/Ud; Modal Repertuvar
West Africa|Blues; Psikedelik/Rock|Epik|Dans; Odak; Spor|Tuareg Çöl Blues'u; Kidal; Tamashek Söz
Western Europe; Latin America; Fusion|Tel/Enstrümantal|Dinginlik|Odak|Enstrümantal Çöl Gitarı; İki Kardeş; Easy Eye Sound
Iranian World; Levant; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak|İran Telleri + Arap Udu ve Perküsyonu
Mediterranean Europe|Caz; Vokal/Şarkı|Melankoli|Gece; Yolculuk|Caz Vokali; Akdeniz ve İtalyan Halk Müziği; Flamenko; Fado; Karayip Etkileri
Western Europe|Vokal/Şarkı|Karanlık|Dans; Yolculuk|İngiliz Halk Şarkısı; Deneysel Folk; Saha Kaydı
Eastern Europe; Western Europe; Fusion|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Polish Folk; Carpatho-Rusyn; North Macedonian; Ukrainian Song
Western Europe|Psikedelik/Rock|Karanlık|Yolculuk|Arabian Fuzz; Psikedelik Rock; Maghreb Füzyon
Western Europe|Psikedelik/Rock|Karanlık|Dans; Odak|Arabian Fuzz; Psikedelik Rock; Maghreb Füzyon
Western Europe|Psikedelik/Rock|Karanlık|Odak; Spor|Arabian Fuzz; Psikedelik Rock; Maghreb Füzyon
Levant|Psikedelik/Rock|Neşe|Dans|Doğu Surf; Psikedelik Rock; Enstrümantal
Levant|Psikedelik/Rock|Neşe|Odak|Psychedelic Rock; Akdeniz Surf; Doğu Rock
East & Central Asia|Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Elektronik Rock/Tekno; Canlı Bateri; Seul
Latin America|Psikedelik/Rock; Hip-Hop/Rap|Neşe|Dans; Spor|Brezilya Rock + Rap
Western Europe|Vokal/Şarkı; Caz|Karanlık|Gece|Fransız Deneysel Sahne; Çoklu Vokal Kadrosu
Western Europe|Caz|Dinginlik|Dans; Odak; Yolculuk|Ethio-Jazz; Psychedelic Jazz; Guest Vocal
West Africa; Latin America; Fusion|Funk/Soul; Gelenek|Neşe|Dans; Odak|Afrobeat Davulu + Kolombiya Ritimleri
Oceania; Western Europe|Pop/Şarkı-Yazarı; Vokal/Şarkı|Dinginlik|Gece|Berlin Sahnesi; Vokal Topluluğu Deneyi
North America|Elektronik|Neşe|Dans; Spor|Ortadoğu Tribal; Ritual Elektronik; Dans Müziği
North America|Elektronik|Epik|Dans; Odak|Ortadoğu Tribal; Ritual Elektronik; Dans Müziği
West Africa|Blues; Psikedelik/Rock|Dinginlik|Gece|Tuareg Gitarı; Ishumar/Assouf; Agadez
North America; Latin America|Caz; Tel/Enstrümantal|Dinginlik|Dans; Odak; Gece; Yolculuk|Contemporary Jazz Quartet; Guitar; Organ; Improvisation
Latin America; Western Europe|Gelenek|Dinginlik|Dans; Odak|And Halk Müziği; Charango/Quena; Paris Yerleşik
Levant; Fusion|Elektronik; Funk/Soul|Neşe|Dans; Odak|Orta Doğu Synth; Funk; Disko; Reggae; Küresel Groove
Western Europe|Funk/Soul; Caz|Neşe|Dans; Odak|Deep Funk / Tru Thoughts Hattı
Caribbean|Gelenek; Vokal/Şarkı|Romantik|Dans; Gece|Kuba Trio Geleneği; Casa Marina Kayıtları
North Africa|Tel/Enstrümantal; Caz|Melankoli|Odak; Gece|Ud; Tunus Medine Repertuvarı; Caz/Etnomüzikoloji; Paris
Northern Europe|Pop/Şarkı-Yazarı; Elektronik|Neşe|Dans; Spor|Trip-Hop; Torch Song; Stockholm
Western Europe|Tel/Enstrümantal; Caz|Melankoli|Odak; Gece|Solo Kontrbas; Sinematik Enstrümantal
Latin America|Elektronik; Gelenek|Dinginlik|Yolculuk|And Elektroniği; Dijital Cumbia
Western Europe|Pop/Şarkı-Yazarı|Neşe|Odak|İrlanda Şarkı-Yazarlığı; Sakin Vokal; Dublin
Latin America|Gelenek; Vokal/Şarkı|Melankoli|Gece|Paraguay Folkloru; Arp/Gitar; Guarani Söz
Latin America|Elektronik; Gelenek|Neşe|Dans; Odak; Spor|Çağdaş Cumbia; Dub/Bass; Bogotá
Western Europe; Balkans; Fusion|Caz; Gelenek|Neşe|Dans; Yolculuk; Spor|Gypsy-Jazz/Balkan Folk; Flüt; Bristol
Balkans|Gelenek; Psikedelik/Rock|Neşe|Dans|Hırvat Halk Ezgileri + Rock; Mojmir Novaković
Western Europe|Gelenek; Klasik/Modern Klasik|Neşe|Dans|Alpine Folk; Yodel; Strings/Accordion
Western Europe; Latin America; Fusion|Elektronik|Karanlık|Dans; Odak; Yolculuk|Elektro-Tango; Downtempo
Western Europe; Balkans; Fusion|Nefesli/Bando; Gelenek|Neşe|Dans; Spor|Balkan Bando + Fransız Sahne Müziği
West Africa|Vokal/Şarkı|Neşe|Dans; Yolculuk|Mbalax; Senegal Griot Geleneği; Hint-Arap Etkisi
Iberia|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Flamenko Cantesi; Jerez Ekolü
North Africa|Elektronik; Gelenek|Neşe|Odak|808 Elektroniği + Mağrip Vokali; Mezoued Hattı
Western Europe|Gelenek|Neşe|Yolculuk|Kelt Folk; Hollanda; Festival İcrası
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Indie Pop; Electro-Singer-Songwriter
North America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Melankoli|Dans; Odak|İnuktitut Vokal; Nunavik; Yerli Halk Anlatısı
Western Europe|Elektronik; Klasik/Modern Klasik|Dinginlik|Dans; Odak; Yolculuk|Progressive Electronica; Field Recordings; Saxophone; Travel-Inspired Composition
Eastern Europe; Caucasus; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Duduk; Türk-Ermeni Repertuvar; Kamara Düzenleme
Iranian World; Northern Europe; Fusion|Vokal/Şarkı|Melankoli|Gece|Afgan Vokali + İskandinav Düzenleme
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Yolculuk|Kolombiya Halk Şarkısı; Toplumsal Söz; Tiple/Gitar; Colombian Protest Song; Folk; Experimental Latin
Eastern Europe|Vokal/Şarkı|Melankoli|Yolculuk|Roman Repertuvarının Slovak Yorumcusu; 'Manuşa' (2022)
Western Europe|Gelenek; Caz|Neşe|Dans; Odak|Klezmer Klarneti; 'Yeni Yahudi Müziği'; Paris
Levant|Caz|Neşe|Dans; Odak; Yolculuk|Çeyrek Tonlu Trompet; Arap Makamı + Caz
Oceania; North America|Pop/Şarkı-Yazarı; Klasik/Modern Klasik|Melankoli|Odak; Spor|Profesyonel İslık; Exotica/Lounge Estetigi; Los Angeles
North Africa|Gelenek|Neşe|Dans; Odak; Spor|Arap Perküsyonu; Darbuka; Oryantal Repertuvar
Western Europe|Gelenek|Melankoli|Gece|Klezmer; Roman Müziği; Yidiş Füzyon
Western Europe; East & Central Asia|Caz|Neşe|Dans; Odak|Saksofon; 'Dirty Jazz'; Paris
Western Europe; Balkans; Fusion|Vokal/Şarkı; Caz|Melankoli|Gece|Balkan Song; Sephardic/Ottoman Repertoire; Piano Trio
Western Europe; South Asia; Fusion|Psikedelik/Rock; Gelenek|Romantik|Dans; Spor|Hindustani Klasik Raga; Psikedelik Rock; Sufi Şiiri; Elektriklendirilmiş Halk Çalgıları
Western Europe; West Africa; Fusion|Klasik/Modern Klasik|Dinginlik|Odak; Gece|Kora-Gitar Düeti; Akdeniz Folk; Kamara Füzyon
Western Europe; West Africa; Fusion|Caz; Gelenek|Neşe|Dans; Odak; Spor|Manding Jazz; Kora; Tenor Saxophone
Southern Africa; Western Europe; West Africa; Fusion|Klasik/Modern Klasik; Gelenek|Melankoli|Odak; Yolculuk|Cello; West African Percussion; Contemporary Chamber
Western Europe; West Africa; Fusion|Vokal/Şarkı; Caz|Melankoli|Gece; Yolculuk|Mandé Geleneği; Protest Şarkı; Gitar-Vokal Düeti
Caucasus|Gelenek; Klasik/Modern Klasik|Melankoli|Gece|Armenian Sacred and Folk; Gurdjieff Repertoire
West Africa|Gelenek; Vokal/Şarkı|Dinginlik|Dans; Odak; Gece|Mande Griot Üçlüsü; Balafon/Ngoni; Kronos İşbirliği
Central Africa|Psikedelik/Rock; Gelenek|Neşe|Yolculuk; Spor|Hurda Enstrüman; Kinşasa; DIY Elektrikli Groove
Latin America|Psikedelik/Rock|Neşe|Dans; Odak|Tropikal Psikedeli; Deneysel Cumbia; Bogotá
Western Europe; Latin America; Fusion|Psikedelik/Rock; Gelenek|Neşe|Dans; Odak; Yolculuk; Spor|Afro-Kolombiya Psikedelyası; Cumbia; Champeta; Karayip Ritimleri; 60'lar Psikedelyası
Latin America|Caz; Funk/Soul|Neşe|Odak; Yolculuk; Spor|Afro-Venezuela Ritimleri + Caz; Bas
Mediterranean Europe|Vokal/Şarkı; Gelenek|Neşe|Dans|Korsika Polifonisi; Paghjella Hattı
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Karanlık|Dans; Yolculuk; Spor|İrlanda Folk + Rock; Şarkı-Yazarlığı
Western Europe; Northern Europe|Gelenek|Neşe|Spor|Franken Halk Müziği + Fin Tangosu
Iranian World|Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Fars Vokali + Flamenko; Lorca Repertuvarı
South Asia|Tel/Enstrümantal; Klasik/Modern Klasik|Epik|Odak; Yolculuk; Spor|Sitar + Sarod + Perküsyon; Hindustani Çağdaşı
Oceania|Vokal/Şarkı; Gelenek|Karanlık|Dans; Spor|Yeni Zelanda Alt-Country; Gotik Folk
Arabian Peninsula; North America; Fusion|Caz; Gelenek|Epik|Dans; Odak; Yolculuk; Spor|Körfez İnci Dalgıcı Müziği + Afro-Küba Cazı
Western Europe|Psikedelik/Rock; Caz|Melankoli|Yolculuk|Deneysel Rock + Çöl Bluesu Gitarı
Caribbean|Reggae/Dub|Karanlık|Odak; Yolculuk|Dub; Roots Reggae; Stüdyo Prodüksiyonu
Iberia; Latin America|Vokal/Şarkı; Reggae/Dub|Neşe|Dans; Spor|Barselona Mestizo Sahnesi; Mestizaje
East & Central Asia; Southeast Asia; Fusion|Gelenek; Vokal/Şarkı|Dinginlik|Gece; Yolculuk|Avustronezya Hattı; Sape + Yerli Vokal
Western Europe|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans|Bluegrass; Welsh Folk; String Band
North America|Vokal/Şarkı; Gelenek|Epik|Yolculuk|Bulgar Kadın Polifonisi
Iranian World|Funk/Soul; Psikedelik/Rock|Neşe|Dans; Spor|Fars Funk-Rock; Tahran Bağımsız Sahne
Iranian World|Psikedelik/Rock|Melankoli|Yolculuk|İran Rock'ı; Hâfız Şiiri; Yeraltı Sahnesi
Iranian World|Gelenek; Tel/Enstrümantal|Neşe|Dans; Spor|Buşehr Afro-Fars Folk; Körfez Etkisi
Iranian World|Pop/Şarkı-Yazarı|Melankoli|Dans|Oda-Pop; Tahran; Farsça Söz
West Africa|Blues; Psikedelik/Rock|Neşe|Dans; Spor|Çöl Blues; Timbuktu'dan Sürgün; Bamako
Iranian World|Pop/Şarkı-Yazarı|Neşe|Yolculuk|Fars Pop; Reggae/Ska Etkisi; Diaspora
Iranian World|Psikedelik/Rock|Neşe|Dans; Spor|Fars Alternatif Rock; 70'ler Pop Cover'ları; Tahran Bağımsız Sahne
East Africa|Funk/Soul|Neşe|Yolculuk|Ethio-Groove; Addis Stüdyo Kadrosu; 70'ler
West Africa|Tel/Enstrümantal; Blues|Neşe|Dans; Yolculuk; Spor|Ngoni; Griot Geleneği; Afro-Blues
Western Europe; East Africa; Fusion|Psikedelik/Rock|Karanlık|Dans; Odak; Yolculuk|Ethio-Rock; Deneysel Etiyopya Müziği; Noise
East Africa|Gelenek; Vokal/Şarkı|Melankoli|Dans|Maloya; Réunion Creole Song; Polyphonic Vocals; Percussion
Latin America|Gelenek|Romantik|Yolculuk|Son Jarocho; Jarana-Charango; Balkan Etkisi
Latin America|Pop/Şarkı-Yazarı; Vokal/Şarkı|Romantik|Dans; Spor|Meksika Halk Şarkısı; Son Jarocho; Bolero
Western Europe; Anatolia; Fusion|Elektronik; Gelenek|Karanlık|Dans; Odak; Yolculuk; Spor|Asit Techno + Anadolu Sazı
North America|Vokal/Şarkı|Melankoli|Gece|Amerikana; Bluegrass Vokal Üçlüsü; Çok Sesli Şarkı
North America|Gelenek; Vokal/Şarkı|Epik|Spor|Indigenous Drum Song; Powwow Singing
Northern Europe|Gelenek; Caz|Melankoli|Gece; Yolculuk|Norwegian Folk; Hardanger Fiddle; Contemporary Jazz; Electroacoustic; Chamber Ensemble
Northern Europe; Eastern Europe; Fusion|Caz|Melankoli|Dans|İskandinav-Slav Füzyon; Arkaik Vokal; Caz
Latin America|Klasik/Modern Klasik|Romantik|Gece; Yolculuk|Tango Cancion; Klasik Tango Vokali
West Africa; Western Europe; Fusion|Tel/Enstrümantal|Dinginlik|Odak; Yolculuk|Kora-Diyatonik Akordeon Düeti; Griot Geleneği
East & Central Asia; South Asia; Fusion|Tel/Enstrümantal|Dinginlik|Odak|Rubab/Sarod + Japon Ezgileri; Paris Yerleşik
Western Europe; Iranian World; Fusion|Caz; Vokal/Şarkı|Melankoli|Yolculuk|Caz Vokali; Peştuca Söz; Köln
West Africa; Western Europe; Fusion|Tel/Enstrümantal; Vokal/Şarkı|Dinginlik|Gece; Yolculuk|Kora + Xalam + Keman; Tamala Topluluğu
Western Europe; West Africa; Fusion|Gelenek; Caz|Neşe|Dans; Odak; Spor|Gnawa; Accordion; West African Jazz
Iranian World|Gelenek|Dinginlik|Odak; Gece; Yolculuk|Fars Klasik Üçlü
Iranian World; Iberia; Fusion|Tel/Enstrümantal; Gelenek|Epik|Odak; Yolculuk|Tar/Setar + Akdeniz Telleri; Modal Müzik
Western Europe; Levant; Fusion|Elektronik; Vokal/Şarkı|Karanlık|Dans; Yolculuk; Spor|Asit Techno + Arapça Vokal
Mediterranean Europe|Elektronik; Caz|Dinginlik|Odak|İtalyan Kolektif Elektroniği; C'mon Tigre Çevresi
Western Europe; North America|Caz; Nefesli/Bando|Dinginlik|Odak; Yolculuk|Trompet + Akordeon Düeti
Latin America|Gelenek; Vokal/Şarkı|Dinginlik|Gece; Yolculuk|And Repertuvarı; Keçuva Dili
North America|Blues|Neşe|Dans; Yolculuk|Blues + Karayip/Afrika/Hint Bağlantıları; Çoklu Enstrüman
Iberia|Klasik/Modern Klasik|Dinginlik|Odak|Nyckelharpa; Sefarad Repertuvarı; Erken Müzik
North Africa|Vokal/Şarkı; Gelenek|Melankoli|Dans; Gece; Yolculuk|Amazigh (Berberi) Şarkı; Taşelhit/Souss; Gitar-Vokal Düo
North America|Gelenek; Blues|Neşe|Dans|Jug Band Canlanması; Eski Usul Amerikan
North Africa|Psikedelik/Rock|Neşe|Dans; Spor|Kabil (Amazigh) Rock; 70'ler; Elektrikli Gitar
Eastern Europe|Funk/Soul; Caz|Neşe|Dans; Yolculuk; Spor|Afrobeat; Jazz-Funk; Moroccan Influences; Horn Section; Multinational Ensemble
North America; Western Europe; Fusion|Blues; Gelenek|Melankoli|Yolculuk|Üç Gelenek Eşit Ağırlıkta: Appalachian/Old-Time, İrlanda Geleneği, Amerikan Ragtime
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Anti-Folk; Akordeon; Viyana
North Africa; Western Europe|Elektronik; Vokal/Şarkı|Melankoli|Dans|Raï/Chaâbi + Kulüp Elektroniği; Lozan
Western Europe|Gelenek; Klasik/Modern Klasik|Dinginlik|Gece; Yolculuk|Alt-Folk; Oda Müziği Enstrümanları; Doğaçlama; Dört Sesli Vokal
Iberia|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Gece|Katalan Şarkı-Yazarı; Kırsal Folk; Ev/Çiftlik Kaydı; Lluçanès
Western Europe|Gelenek; Nefesli/Bando|Dinginlik|Odak|İskoç ve İrlanda Geleneksel Müziği; Boehm Flüt; Whistle; Keman; Geleneksel Ezgi Yazımı
North Africa|Elektronik; Caz|Karanlık|Odak; Yolculuk|Deneysel; Mikrotonal Topluluğu; Kahire Bağımsız Sahnesi
North Africa|Elektronik; Caz|Neşe|Dans; Odak|Deneysel; Mikrotonal Topluluğu; Kahire Bağımsız Sahnesi
Iranian World|Vokal/Şarkı|Dinginlik|Gece|İran Bölgesel Halk Müziği; Çok Dilli Vokal
Anatolia; Western Europe|Tel/Enstrümantal; Caz|Dinginlik|Odak; Yolculuk|Perdesiz Ud; Anadolu Makamı + Caz; Amsterdam
Anatolia; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak|Bağlama + Perküsyon; Çok Uluslu Kadro
North Africa|Vokal/Şarkı|Epik|Yolculuk; Spor|Moritanya Griot Geleneği; Ardin; Sahra Psikedelisi
Western Europe|Psikedelik/Rock|Neşe|Dans; Spor|İngiliz Folk-Rock; Gitar; Şarkı Yazımı
Balkans|Tel/Enstrümantal; Gelenek|Melankoli|Odak; Gece; Yolculuk|Buzuki; Yunan Repertuvarı (Theodorakis Bestesi)
Mediterranean Europe|Pop/Şarkı-Yazarı; Gelenek|Neşe|Dans; Spor|Napoli Lehcesi; Elektronik + Halk Ezgisi
North Africa|Tel/Enstrümantal|Dinginlik|Odak|Ud; Endelüs/Mağrip Repertuvarı; 3MA
Anatolia; Iranian World; Fusion|Gelenek|Dinginlik|Gece; Yolculuk|Anadolu-Fars Kamara; Bağlama; Kadın Vokal
North Africa|Caz; Tel/Enstrümantal|Epik|Odak; Yolculuk; Spor|Ud; Sufi Vokal; Çağdaş Caz
Balkans|Gelenek|Neşe|Dans|Bulgar Halk Müziği; Didgeridoo-Kaval; Dubstep-Drum'n'Bass
Anatolia|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Karadeniz/Laz Repertuvarı; Kemençe Hattı
Iberia; Anatolia; Fusion|Gelenek|Melankoli|Gece; Yolculuk|Klezmer; Balkan Füzyon; Sefarad Repertuvarı
Anatolia|Vokal/Şarkı|Melankoli|Yolculuk|Kürtçe Şarkı; Anadolu Vokal Geleneği; Çağdaş Düzenleme
Western Europe|Pop/Şarkı-Yazarı; Reggae/Dub|Neşe|Dans; Yolculuk|Çağdaş Chanson; Reggae/Ska; Toplumsal Söz
Balkans|Vokal/Şarkı; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Yunan Şarkı-Yazarı; Arapça/Doğu Akdeniz Repertuvarı
Oceania; Iranian World; South Asia; Fusion|Psikedelik/Rock; Gelenek|Karanlık|Odak; Yolculuk|Progressive Rock; Persian Setar; Afghan Rabab; Drone
Iberia|Vokal/Şarkı|Epik|Gece; Yolculuk|Minimal Vokal-Gitar; Katalan Şarkısı; Deneysel Prodüksiyon
Western Europe|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|İngiliz Geleneksel Folk; Folk Revival; Parmak Stili Gitar
Iberia|Tel/Enstrümantal; Gelenek|Neşe|Odak|Flamenko Gitar; Toque; Roman Virtüozitesi
North America|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Bluegrass; Americana; Dobro
Western Europe|Gelenek; Elektronik|Melankoli|Dans; Yolculuk; Spor|İrlanda Deneysel Folk; Geleneksel Şarkı; Drone/Ambient; Elektronik
Oceania; Iranian World; Fusion|Caz; Gelenek|Dinginlik|Dans; Odak; Gece; Yolculuk|Pers Caz Füzyonu; Tar; Oda Müziği Doğaçlaması
Iranian World; Oceania; Fusion|Caz; Gelenek|Melankoli|Odak; Gece; Yolculuk|Persian Tar; Contemporary Jazz; Ensemble Improvisation
East & Central Asia|Caz; Vokal/Şarkı|Dinginlik|Odak; Gece|Mongol Caz; Urtiin Duu Etkisi; Oda Cazı
Western Europe; Levant; Fusion|Caz; Nefesli/Bando|Dinginlik|Odak; Gece; Yolculuk|Arap-Caz Füzyonu; Klezmer Klarnet; Nefes Müziği
Iberia; Levant; Fusion|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk|Sefarad; Ladino; Osmanlı-Balkan Repertuvarı
Western Europe|Caz|Neşe|Dans; Odak; Spor|Gypsy Jazz; Manouche; Keman-Gitar Swing
Caribbean; North America; Fusion|Hip-Hop/Rap; Gelenek|Neşe|Dans|Haiti Rap; Rasin; Afro-Karayip Füzyon
Iranian World; Western Europe; Fusion|Klasik/Modern Klasik; Gelenek|Romantik|Odak|Fars Radifi; Arap Makamı; Tar; Kanun; Viola da Gamba; Akdenizlerarası Oda Müziği
Eastern Europe|Gelenek; Vokal/Şarkı|Neşe|Dans|Macar Roman Müziği; Oláh Cigány; Ağız Bası Perküsyon
Western Europe; Latin America; Fusion|Gelenek; Nefesli/Bando|Dinginlik|Dans|Pan-Latin Halk Müziği; And Müziği; Quena/Siku; Charango; Arp; Pan-Latin Repertuvar
West Africa|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Moor Müziği; Ardîn-Tidinit; Sahel Blues Etkisi
Mediterranean Europe|Gelenek|Epik|Dans; Gece; Yolculuk|Tarantella; Balkan-Akdeniz Füzyon; Tango Etkisi
Mediterranean Europe; Fusion|Gelenek; Caz|Neşe|Dans; Odak|Rumba Flamenca; Manouche; Akdeniz Halk Müziği
Iberia; Fusion|Elektronik; Gelenek|Dinginlik|Odak|Oriental Elektronik; Ud-Ney-Klarnet; Organik Perküsyon
Oceania; Mediterranean Europe; Fusion|Tel/Enstrümantal; Klasik/Modern Klasik|Epik|Odak; Gece; Yolculuk|İspanyol/Flamenko Gitarı; Klasik Gitar; Latin Akustik Gitar
North Africa|Klasik/Modern Klasik|Epik|Odak; Gece|Film Müziği; Orkestral Arap; Sahel Renkleri
Central Africa|Psikedelik/Rock; Elektronik|Neşe|Dans; Spor|Eko-Punk; Geri Dönüşüm Perküsyon; Kinşasa Avangard
East Africa; North America; Fusion|Gelenek; Psikedelik/Rock|Melankoli|Yolculuk|Ruanda Halk Müziği; Deneysel Gitar; Minimal Akustik
Northern Europe; East Africa; Fusion|Psikedelik/Rock; Hip-Hop/Rap|Epik|Dans; Spor|Deneysel Psychedelia; Luganda Rap; Kabile Ritmi Kurgusu
Latin America; Western Europe; Fusion|Elektronik|Karanlık|Dans|Deneysel Elektronik; Ambient Art Pop; Bolero Kalıntıları
North America|Gelenek; Vokal/Şarkı|Dinginlik|Yolculuk|Old-Time Banjo; Appalachian Repertuvar; Clawhammer
North America|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Dans; Spor|Bluegrass; Flatpicking Gitar; Çağdaş Americana
North America|Blues; Vokal/Şarkı|Melankoli|Spor|Rhythm and Blues; Garage Gospel; New Orleans Geleneği
North America|Pop/Şarkı-Yazarı; Gelenek|Neşe|Spor|Americana; Country-Folk; Nashville Sahnesi
Levant|Psikedelik/Rock|Romantik|Spor|Psychedelic Rock; Akdeniz Surf; Doğu Rock
Latin America; Western Europe; Fusion|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Dans; Gece; Yolculuk|Afro-Venezuela Perküsyonu; Barlovento Ritüeli; Barok-Halk Füzyonu
Western Europe|Pop/Şarkı-Yazarı|Karanlık|Gece|Arp Eşlikli Indie; Dream Pop; Newcastle Sahnesi
Western Europe; Anatolia; Fusion|Elektronik|Melankoli|Yolculuk|Balkan Elektro; Bucovina Club; Frankfurt
Northern Europe; East Africa; West Africa; Fusion|Caz; Funk/Soul|Karanlık|Yolculuk; Spor|Ethio-Jazz; Mali Blues; Afrobeat; Scandinavian Jazz; Psychedelic Soul
Western Europe|Tel/Enstrümantal|Karanlık|Odak|American Primitive Gitar; Clawhammer Banjo; Gal Kökeni
West Africa; Western Europe; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece|Kora-Trompet İkilisi; Griot Geleneği; Tango Etkisi
Western Europe; Latin America; Fusion|Elektronik; Pop/Şarkı-Yazarı|Melankoli|Dans|Africa Express Kolektifi; Sınır-Aşırı İşbirliği; Bahidorá Kaydı
Latin America|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Currulao; Marimba de Chonta; Pasifik Kolombiya
Levant|Caz; Klasik/Modern Klasik|Dinginlik|Dans; Odak|Çağdaş Caz; Orta Doğu Makamları; Çeyrek Sesli Klavyeler; Progresif Rock/Metal; Elektronika
Levant; Western Europe; West Africa; Fusion|Caz; Funk/Soul|Epik|Dans; Odak; Yolculuk; Spor|Kontrbas Deneyi; Blufunk; Serbest Doğaçlama
North America; Eastern Europe; Fusion|Gelenek; Tel/Enstrümantal|Dinginlik|Odak; Gece; Yolculuk|Klezmer Keman; Nigun; Yidiş Repertuvar
East & Central Asia|Vokal/Şarkı; Gelenek|Neşe|Dans; Spor|Kore Minyo; Gyeonggi Sori; Çağdaş Düzenleme
Iranian World; Western Europe|Gelenek; Tel/Enstrümantal|Melankoli|Odak; Gece|Setar/Daf + Zarb; Fars Modal Geleneği
Levant|Vokal/Şarkı|Melankoli|Dans|Arap Şarkı Geleneği; Rahbani Besteleri; Beyrut
Iberia|Vokal/Şarkı; Gelenek|Romantik|Gece|Copla; Flamenko; Endülüs Şarkı Geleneği
West Africa|Gelenek; Caz|Neşe|Odak; Gece; Yolculuk|Mande Griot; Balafon-Gitar Diyaloğu; Afrika Blues
East Africa|Caz|Melankoli|Odak; Gece|Ethio-Jazz'ın Kurucusu; Vibrafon; Addis Ababa
Iberia|Klasik/Modern Klasik; Gelenek|Epik|Gece|Concertina Kuarteti; Portekiz Oda Müziği; Diyatonik Akordeon
Western Europe|Gelenek; Elektronik|Dinginlik|Gece; Yolculuk|Gal Üçlü Arpı; Deneysel Folk; Cymraeg Repertuvar
Northern Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Gece|Norveç Halk Müziği; Nord-Østerdalen Geleneği; Halk Şarkıcılığı/Kveding; Gitar; Bukkehorn; Seljefløyte; Lir
Levant; Iberia; Fusion|Gelenek; Tel/Enstrümantal|Romantik|Odak|Ud-Gitar İkilisi; Maqam Iraqi; Endülüs Bağı
North Africa; Western Europe; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Yolculuk|Flamenko-Arap Füzyon; Ud; Endülüs Mirası
Western Europe|Vokal/Şarkı; Caz|Melankoli|Gece|Retro Şansonet; Vokal Caz; 1930'lar Estetigi
Western Europe; Fusion|Elektronik; Funk/Soul|Dinginlik|Dans; Odak; Yolculuk; Spor|Psychedelic Dub; Afro-Latin Ritim; Gnawa-Cumbia Füzyonu
Mediterranean Europe; Fusion|Funk/Soul; Elektronik|Neşe|Dans; Odak; Yolculuk|Cumbia; Afro-Funk; Salento; Vokalsüz Enstrümantal
North Africa; North America; Fusion|Psikedelik/Rock; Elektronik|Karanlık|Odak|Kahire Deneysel; Serbest Doğaçlama; Agouza Sahnesi
Oceania; Balkans; Mediterranean Europe; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece; Yolculuk|Balkan/Akdeniz Dünya Cazı; Ud; Kontrbas; Perküsyon; Piyano; Doğaçlama
North America|Blues; Funk/Soul|Neşe|Dans; Spor|Elektrik Blues; Soul-Blues; Slide Gitar
Western Europe; Anatolia; Fusion|Vokal/Şarkı; Psikedelik/Rock|Epik|Dans; Spor|Zazaki Şarkı; Kürt Şiiri; Caz-Rock Doku
Western Europe|Funk/Soul; Caz|Dinginlik|Dans; Yolculuk|Afro-Diaspora Füzyonu; Psikedelik Trip-Hop; Kozmik Neo-Soul; Dub Dokular
North Africa; Western Europe|Vokal/Şarkı; Gelenek|Melankoli|Dans|Endülüs-Arap Geleneği; Gnawa Kökleri; Folk-Pop; Cezayir Darija Vokali
Western Europe; Mediterranean Europe|Vokal/Şarkı; Klasik/Modern Klasik|Melankoli|Gece|Çok Dilli Şarkı; Oda Müziği Folk; Bossa ve Flamenko Etkisi; İrlanda Folk Dokunuşu
Mediterranean Europe; Balkans; Fusion|Vokal/Şarkı; Klasik/Modern Klasik|Karanlık|Yolculuk|Yunan Vokal; Serbest Doğaçlama; Akdeniz Oda Müziği
Balkans; Mediterranean Europe; Fusion|Elektronik; Gelenek|Romantik|Dans; Yolculuk|Sevdah Remix; Balkan Elektronik; Disko Edit
Anatolia|Funk/Soul; Gelenek|Neşe|Dans; Odak; Spor|Anadolu Funk; Oryantal Soul; 70'ler Kayıt Estetigi
Levant|Gelenek; Tel/Enstrümantal|Romantik|Odak; Gece; Yolculuk|Buzuq; Gitar/Elektronik; Arap Halk Müziğinin Yeniden Yorumu; Deneysel Oda Müziği
Levant|Psikedelik/Rock|Neşe|Odak; Spor|Surf Rock; Akdeniz Gitar Müziği; Kütüphane Müziği; Levanten Break; Kıyı Psikedelisi; Vokalsüz Enstrümantal
East & Central Asia|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Sani/Yi Halk Müziği; Yunnan; Kolektif Şarkı
Latin America|Gelenek; Vokal/Şarkı|Neşe|Dans|Carimbó; Amazon Yerli Müziği; Nheengatu/Portekizce; Kadın Kolektifi
Northern Europe|Caz; Gelenek|Melankoli|Yolculuk; Spor|Fin Dünya Füzyonu; Etno-Caz; Saksafon-Kantele
Mediterranean Europe|Vokal/Şarkı; Gelenek|Epik|Dans|Ritüel/Trans Dünya Müziği; Çerçeve Davulları; Vokal; Akustik Gitar; Şamanik/Koral
West Africa|Elektronik; Hip-Hop/Rap|Neşe|Dans; Yolculuk; Spor|Lo-Fi Highlife; Erken Hiplife; Kaset Estetigi
North Africa; West Africa; North America; Fusion|Caz; Gelenek|Dinginlik|Odak; Gece|Ud-Kora-Bateri; Serbest Doğaçlama; Kültürlerarası Caz
North Africa|Vokal/Şarkı|Epik|Spor|Moritanya Griot Geleneği; Ardin; Sahra Psikedelisi
Latin America; Western Europe; Fusion|Gelenek; Caz|Neşe|Dans; Spor|Tango Negro; Afro-Arjantin Kökler; Candombe
Balkans|Klasik/Modern Klasik; Elektronik|Neşe|Dans; Odak; Spor|Film Müziği; Yunan Sineması; Atmosferik Doku
Western Europe|Elektronik; Klasik/Modern Klasik|Dinginlik|Odak; Yolculuk; Spor|Çello-Gitar-Bateri Üçlüsü; Elektronik Avangard; Oda Müziği + Dünya Etkisi
Western Europe|Caz; Klasik/Modern Klasik|Neşe|Yolculuk|Barok-Caz Füzyonu; Stride Piyano; Berlin Filarmoni Serisi
Levant|Gelenek; Tel/Enstrümantal|Neşe|Dans; Odak; Yolculuk; Spor|Kanun; Filistin Tarab; Çağdaş Arap Doğaçlaması
Latin America|Gelenek|Romantik|Gece|Akustik Folk; Latin Songwriter
Levant; Western Europe; Fusion|Gelenek; Caz|Dinginlik|Dans; Odak; Gece; Yolculuk|Ud; Hicaz Makamı; Arap Oda Müziği
Western Europe|Gelenek; Vokal/Şarkı|Dinginlik|Odak; Gece|Geleneksel Ezgi; Akordeon-Kontrbas İkilisi; Tango Etkisi; Sözsüz Vokal
Western Europe|Vokal/Şarkı; Gelenek|Romantik|Gece|Gal Halk Şarkısı; Cymraeg Repertuvar; Akustik Folk
Western Europe|Caz|Epik|Odak; Spor|Akordeon Cazı; Living Being Projesi; Avrupa Çağdaş Cazı
Levant|Pop/Şarkı-Yazarı; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Arap Pop; Mısır-Lübnan Prodüksiyon; Anaakım
Western Europe; Iranian World; Balkans; Levant; Fusion|Caz; Gelenek|Neşe|Dans; Odak; Yolculuk|Zarb + Çello + Piyano; Çok Gelenekli Oda Müziği
South Asia|Gelenek; Vokal/Şarkı|Neşe|Dans; Yolculuk; Spor|Manganiyar Geleneği; Thar Çölü; Sindhi Halk Müziği
North America; West Africa; Fusion|Psikedelik/Rock; Gelenek|Epik|Odak; Yolculuk; Spor|Mali Vokali + Psikedelik Rock; Oakland
Southeast Asia; Fusion|Psikedelik/Rock; Funk/Soul|Epik|Odak|Tropikal Çöl Müziği; Keroncong; Dangdut; Orta Doğu Groove'ları; Afro-Funk; Enstrümantal
Balkans; Fusion|Reggae/Dub; Nefesli/Bando|Neşe|Dans; Spor|Balkan Dub; Gypsy Brass; Bulgar Koro; Asphalt Tango
Western Europe|Elektronik|Neşe|Dans; Odak; Spor|Enstrümantal Elektronik; Asya-Ortadoğu Sample; Psikedelik
Latin America|Tel/Enstrümantal; Caz|Melankoli|Odak|Bossa Nova; Afro-Samba; Klasik Gitar Tekniği
North America|Psikedelik/Rock|Karanlık|Dans|Gothic Americana; Alt-Country; Appalachian Revival
North America|Psikedelik/Rock|Karanlık|Dans; Odak; Spor|Gothic Americana; Alt-Country; Appalachian Revival
Latin America; Western Europe; Fusion|Klasik/Modern Klasik; Vokal/Şarkı|Epik|Odak; Gece; Yolculuk|Klarnet-Çello İkilisi; Arjantin Folkloru; Bagualá
Western Europe|Pop/Şarkı-Yazarı; Gelenek|Karanlık|Odak|İngiliz Nu-Folk; Şarkı-Yazarı; Açık Akort Gitar
West Africa|Psikedelik/Rock; Funk/Soul|Epik|Odak; Gece; Yolculuk|Nijerya Afro-Rock; 70'ler Funk; Power Trio
Western Europe|Vokal/Şarkı; Caz|Melankoli|Gece|İngiliz Folk-Caz; Echoplex Gitar; Şarkı-Yazarı
North America; Western Europe; Fusion|Caz|Dinginlik|Odak; Spor|Çağdaş Caz; Beat-Odaklı Doğaçlama; Chicago-Londra Ekseni
Western Europe; North America; Fusion|Caz; Psikedelik/Rock|Karanlık|Odak; Yolculuk|Jazz Fusion; Hint Etkisi; Elektrik Gitar Virtüozitesi
Western Europe|Gelenek; Caz|Melankoli|Odak|Klezmer Klarneti; 'Yeni Yahudi Müziği'; Paris
West Africa|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk; Spor|Benin Afro-Funk; Vodun Ritimleri; Albarika Store Kayıtları
Western Europe; South Asia; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Dans|İngiliz Folkü + Hint Klasiği; Sitar-Keman
Northern Europe; Mediterranean Europe; Fusion|Elektronik; Funk/Soul|Neşe|Dans; Odak; Yolculuk; Spor|Afro-Elektronik; Küresel Funk; Afrobeat; Caz Nefeslileri; Poliritmik Prodüksiyon; Elektronik Dans Müziği
Western Europe; Fusion|Caz|Epik|Odak; Yolculuk|Modern Caz; Kuzey Afrika Tonları; Post-Rock Kanatlar; Neo-Klasik Doku; Sözsüz Vokal
South Asia; Western Europe; Fusion|Caz; Gelenek|Melankoli|Dans; Odak|Hint-Caz Füzyonu; Tabla-Bateri; Londra Sahnesi
Latin America|Elektronik|Karanlık|Dans; Yolculuk|Deneysel Elektronik; Other People; New York-Santiago
North America|Pop/Şarkı-Yazarı; Gelenek|Melankoli|Gece|Kanada Folk; Şarkı-Yazarı; Akustik Oda Dokusu
East Africa|Gelenek; Vokal/Şarkı|Dinginlik|Gece; Yolculuk|Ruanda Övgü Şarkısı; Elektrik Gitar; Bas; Çağrı-Yanıt Vokalleri; 1980'ler Stüdyo Folk'u
Iberia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Gece|Flamenko; Katalan Kökeni; Çağdaş Prodüksiyon
West Africa|Blues; Psikedelik/Rock|Melankoli|Dans; Odak; Yolculuk|Tuareg Çöl Blues'u; Kidal; Tamashek Söz
North America|Pop/Şarkı-Yazarı|Karanlık|Gece|Karanlık Folk; Parmak Tezene Akustik Gitar; Viyola Renkleri; Şiirsel Melankoli
Caribbean|Funk/Soul; Gelenek|Neşe|Dans; Odak; Spor|Afro-Küba Ritimleri; Funk Bas; İspanyolca Söz Oyunu; Karayip Mizahı
North America|Gelenek; Pop/Şarkı-Yazarı|Epik|Spor|Appalachian Folk; Old-Time; Holler Repertuvarı
Caribbean|Caz; Gelenek|Epik|Odak; Yolculuk; Spor|Latin Caz; Afro-Karayip Ritim; Cimarrón Teması
Latin America|Pop/Şarkı-Yazarı|Melankoli|Dans|Meksika Pop-Rock; Akordeon; Norteña Etkisi
Northern Europe|Gelenek|Epik|Gece; Yolculuk|Fin Halk Müziği; Sami Joik; Akordeon-Kantele; İklim Aktivizmi
Northern Europe|Gelenek; Tel/Enstrümantal|Dinginlik|Odak|Norveç Slått; Hardingfele; Fele İkilisi
Western Europe|Gelenek; Caz|Neşe|Odak; Gece|Balfolk; Enstrümantal Flaman Folk; Caz ve Funk Etkisi; Dans Müziği
North Africa; Western Europe; Fusion|Gelenek; Klasik/Modern Klasik|Epik|Odak|Ud; Endülüs-Mağrip Repertuvarı; Orkestral Arap
North Africa|Vokal/Şarkı; Elektronik|Neşe|Dans; Spor|Kahire Bağımsız; Arap Alternatif; Şeyh İmam Repertuvarı
Western Europe; Balkans; Iranian World; Fusion|Klasik/Modern Klasik; Gelenek|Melankoli|Odak; Gece; Yolculuk|Trakya Repertuvarı; Çello-Lyra-Zarb; Batı-Doğu Karşılaşması
North America|Gelenek; Blues|Melankoli|Gece|Siyah Telli Çalgı Geleneği; Banjo/Keman; Pulitzer
West Africa|Funk/Soul; Gelenek|Neşe|Dans; Yolculuk; Spor|Benin Afro-Funk; Vodun Ritimleri; 45'lik Kayıtlar
North America|Funk/Soul|Neşe|Dans; Odak; Spor|Latin Soul/Funk; Black Pumas; Austin
North America|Psikedelik/Rock; Gelenek|Karanlık|Dans; Odak|Deneysel Ritüel Folk; Lo-Fi; Çöl Estetigi
North America|Gelenek; Pop/Şarkı-Yazarı|Neşe|Dans; Spor|Americana; Country; Bluegrass
North America|Pop/Şarkı-Yazarı|Melankoli|Gece|Karanlık Folk; Minimalist Folk; Alternatif Country; Keman-Dobro Dokusu; Lo-fi Kayıt
Western Europe; Caribbean; Fusion|Reggae/Dub; Elektronik|Neşe|Dans; Spor|Digital Dub; Dancehall; Remiks Prodüksiyonu
Western Europe|Gelenek; Elektronik|Karanlık|Gece|Northumbrian Folk; Elektronik İşleme; Balad Yeniden Yorumu
Latin America|Vokal/Şarkı; Elektronik|Epik|Gece|Deneysel Vokal; Kolombiya Folkloru; Döngü İşleme
Iberia; Latin America; West Africa; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Dans; Gece|Lizbon Kolektifi; Morna-Bossa Kesişimi; Lusofon Ağ
North America; Latin America; Fusion|Gelenek|Epik|Gece|Nuevo Mexicano; Huapango; Akordeon/Keman
Latin America; Western Europe; Fusion|Vokal/Şarkı; Caz|Neşe|Dans; Odak; Yolculuk; Spor|Tropicália Etkisi; MPB; Londra Caz Sahnesi; Canlı Doğaçlama
North America|Gelenek; Tel/Enstrümantal|Neşe|Odak; Gece|Bluegrass Flatpicking; Dawg Music; Akustik Gitar
North America; Fusion|Gelenek; Tel/Enstrümantal|Neşe|Odak; Gece|Yaylı Topluluk; Göçmen 78'lik Repertuvarı; İtalyan Mandolin; Kalipso ve Venezuela Valsı; Cape Verde Kemanesi
North America|Tel/Enstrümantal; Caz|Neşe|Odak|Oud-Keman; Ortadoğu-Balkan Füzyon; Enstrümantal
Western Europe|Pop/Şarkı-Yazarı|Karanlık|Gece|Samimi Folk; Şarkıcı-Şarkı Yazarı; Şanson; Atmosferik Bedroom Folk
Iberia|Reggae/Dub; Gelenek|Neşe|Dans; Yolculuk|Bask Reggae; Sound System; Euskara Repertuvar
Mediterranean Europe|Gelenek; Vokal/Şarkı|Epik|Dans; Spor|İtalyan Halk Canlanması; Canzone Popolare; Bella Ciao Projesi
Northern Europe|Gelenek; Vokal/Şarkı|Karanlık|Odak; Spor|İsveç Halk Müziği; Queer-Feminist Yorum; Koro Dokusu
Western Europe; East & Central Asia; Fusion|Gelenek; Elektronik|Dinginlik|Spor|Khöömii Gırtlak Şarkısı; Morin Khuur-Tovshuur; Avrupa Elektro Doku
Western Europe; Balkans; Fusion|Caz|Neşe|Odak|Modal Caz; Yunan Ada Renkleri; Karşılama
Western Europe|Gelenek; Tel/Enstrümantal|Melankoli|Spor|İngiliz Folk; 12-Telli Gitar; Peyzaj Anlatısı
Western Europe; North America; Fusion|Gelenek; Pop/Şarkı-Yazarı|Melankoli|Spor|İngiliz Folk Canlanması; Amerikan Indie Yorumu; Balad Repertuvarı
Western Europe|Elektronik|Karanlık|Dans; Odak; Yolculuk; Spor|Trip-Hop; Bristol Sound; Dub Etkisi
Levant|Vokal/Şarkı|Melankoli|Dans; Spor|Arap Şarkı Geleneği; Rahbani Besteleri; Beyrut
Levant|Elektronik|Dinginlik|Gece|Arap Elektronik; Alternatif Pop; Trip-Hop
Western Europe|Gelenek; Caz|Neşe|Dans|İngiliz Folk-Rock; Jansch/Renbourn Gitarları; 60'lar
Anatolia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans; Odak|Türk Hafif Müziği; 45'lik Dönemi; Grafson Kayıtları
Anatolia; Western Europe; Fusion|Gelenek; Caz|Melankoli|Yolculuk|Bağlama; Perküsyon; Kültürlerarası Proje
Anatolia|Gelenek; Vokal/Şarkı|Melankoli|Dans|Çok Sesli Kadın Vokal Topluluğu; Anadolu/Balkan/Kafkas Halk Müziği; Çok Dilli; Geleneksel Çalgılar
Anatolia|Vokal/Şarkı; Gelenek|Melankoli|Dans|Türk Sanat Müziği; Perdesiz Gitar; Çağdaş Düzenleme
Anatolia|Vokal/Şarkı; Gelenek|Neşe|Dans|Türkü Yorumu; Nâzım Hikmet; Odeon Kayıtları
Anatolia|Gelenek; Vokal/Şarkı|Melankoli|Yolculuk|Bozlak; Abdal Geleneği; Kırşehir Bağlaması
Anatolia|Vokal/Şarkı; Gelenek|Romantik|Dans; Spor|Karadeniz Repertuvarı; Kemençe; Türkü Yorumu
Anatolia|Gelenek; Tel/Enstrümantal|Melankoli|Dans; Gece; Yolculuk|Perdesiz Gitar; Kopuz; Aşık Geleneği
Anatolia; Western Europe; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Gece|Anadolu Vokali; Etno-Caz; Diaspora Repertuvarı
Anatolia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Gece; Yolculuk|Geniş Vokal Aralığı; Türkçe Şarkı; Akustik Düzenleme
North America; Anatolia; Fusion|Vokal/Şarkı; Gelenek|Melankoli|Gece|Türkü Yorumu; Balkan Repertuvar; Alan Araştırması
Anatolia|Elektronik; Gelenek|Neşe|Dans; Spor|Anadolu Rock Remiks; Protest Türkü; Elektronik Yeniden Yorum
North America; Anatolia; Fusion|Elektronik; Gelenek|Melankoli|Dans; Odak; Spor|Anadolu Elektronik; Toronto Sahnesi; Türkü Yeniden Düzenleme
Anatolia|Klasik/Modern Klasik; Vokal/Şarkı|Epik|Yolculuk|Çağdaş Klasik; Yunus Emre Bestesi; Piyano-Vokal
Anatolia|Gelenek; Vokal/Şarkı|Neşe|Dans|Kürt Halk Müziği; Dengbêch Geleneği; Kolektif Koro
Anatolia|Gelenek; Vokal/Şarkı|Neşe|Dans; Spor|Kürt Halk Müziği; Dengbêch Geleneği; Kolektif Koro
Anatolia; North America; Fusion|Psikedelik/Rock; Elektronik|Dinginlik|Yolculuk|İstanbul Psychedelia; Oriental Dub; Elektro Saz
Anatolia|Vokal/Şarkı; Pop/Şarkı-Yazarı|Melankoli|Dans; Spor|Protest Şarkı; Arabesk-Halk Kesişimi; Sürgün Repertuvarı
Western Europe|Pop/Şarkı-Yazarı|Melankoli|Gece|Fransız Şansonu; Söz Odaklı Yazarlık; Elektro ve Rock Dokular; Piyano-Saksofon
North America; Anatolia; Fusion|Funk/Soul; Nefesli/Bando|Neşe|Dans; Yolculuk; Spor|Acid Jazz; Balkan-Roman Nefesli; Türk Funk Füzyonu
North America; Anatolia; Fusion|Funk/Soul; Nefesli/Bando|Neşe|Dans; Odak; Yolculuk; Spor|Acid Jazz; Balkan-Roman Nefesli; Türk Funk Füzyonu
Anatolia; North America; Fusion|Caz; Funk/Soul|Neşe|Dans; Yolculuk; Spor|Türk Perküsyonu; Harmolodic Bas; Groove Füzyonu
Iberia|Vokal/Şarkı; Gelenek|Melankoli|Gece; Yolculuk|Valensiya Halk Şarkısı; Katalan Repertuvar; Akdeniz Füzyonu
West Africa|Vokal/Şarkı|Neşe|Dans; Spor|Morna; Coladeira; Cape Verde Geleneği
Levant|Gelenek; Caz|Romantik|Odak; Gece; Yolculuk|Ud; Arap Cazı; Suriye Diasporası
Levant; North America|Vokal/Şarkı; Gelenek|Melankoli|Yolculuk|Yiddish Song; Klezmer; Vocal Ensemble
North Africa|Vokal/Şarkı; Gelenek|Neşe|Dans; Yolculuk; Spor|Chaabi; Cezayir Kabare; Kasbah Repertuvarı
Caucasus; North America; Fusion|Gelenek; Elektronik|Melankoli|Odak; Gece; Yolculuk|Duduk; Ambient Prodüksiyon; Ermeni Ağıt Geleneği
Anatolia|Vokal/Şarkı|Romantik|Dans; Yolculuk; Spor|Rebetiko; İstanbul-İzmir Repertuvarı; Ege Geleneği
West Africa|Gelenek; Funk/Soul|Melankoli|Dans; Yolculuk; Spor|Highlife; Afro-Funk; Aşiko
North Africa|Vokal/Şarkı|Melankoli|Gece|Cezayir Şarkı Geleneği; Folk; Chaâbi Füzyon
West Africa|Funk/Soul|Melankoli|Dans; Yolculuk; Spor|Ghanaian Highlife; Afro-Funk; Dance Band; Guitar Groove
Oceania|Funk/Soul; Caz|Dinginlik|Odak; Yolculuk|Sinematik Soul; Library Müziği Estâtiği; Melbourne Sahnesi
Balkans|Pop/Şarkı-Yazarı|Dinginlik|Gece|Yunan Indie-Folk; İngilizce/Yunanca Söz

/* --- G SÜTUNU BİTTİ --- */
`;

var OW_SYN = `
/* --- I SÜTUNU BURAYA --- */
Caz;Jazz
Ud;Oud
Gitar;Guitar
Piyano;Piano
Perküsyon;Percussion
Doğaçlama;Improvisation
Elektronik;Elektronika;Electronic;Electronica
Flamenko;Flamenco
Çello;Cello
Americana;Amerikana
Keman;Violin
Akordeon;Accordion
Arp;Harp
Klarnet;Clarinet
Flüt;Flute
Nefesli;Brass
Vokal;Vocal;Vocals
Çöl Blues;Desert Blues;Tuareg Rock;Assouf;Tuareg
Geleneksel;Traditional
Füzyon;Fusion
Davul;Drums
Yaylı;Strings
Şarkı;Song
Oda Müziği;Chamber
Topluluk;Ensemble
Ritim;Rhythms
Etki;Influence
Chaabi;Chaâbi
Mande;Mandé

/* --- I SÜTUNU BİTTİ --- */
`;
