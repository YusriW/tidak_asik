import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from collections import defaultdict
import io
import os

def extract_xml_to_excel(xml_content):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    # INFO RESPON
    data = {
        "npwp": root.findtext(".//lembagaJasaKeuangan/npwpLjk"),
        "nama": root.findtext(".//lembagaJasaKeuangan/namaLjk"),
        "no_respon": root.findtext(".//suratJawaban/noRespon"),
        "tgl_respon": root.findtext(".//suratJawaban/tglRespon"),
        "no_surat_permintaan": root.findtext(".//suratJawaban/noSuratPermintaan"),
        "status_respon": root.findtext(".//suratJawaban/statusRespon"),
        "nama_pj": root.findtext(".//suratJawaban/namaPj"),
        "jabatan_pj": root.findtext(".//suratJawaban/jabatanPj")
    }
    df_info_respon = pd.DataFrame([data])

    # DAFTAR WP
    daftar_wp = {
        "npwp": root.findtext(".//responPermintaan/responData/npwp"),
        "nama_wp": root.findtext(".//responPermintaan/responData/namaWp"),
        "NIK": root.findtext(".//responPermintaan/responData/nik", default=''),
        "status_nasabah": root.findtext(".//responPermintaan/responData/statusNasabah"),
        "jml_data": len(root.findall(".//responPermintaan/responData/dataKeuangan/dataRekening"))
    }
    df_daftar_wp = pd.DataFrame([daftar_wp])

    # DAFTAR REKENING
    dict_saldo_awal = {}
    rekening_count = defaultdict(int)
    list_daftar_rekening = []

    for data_rek in root.findall(".//responPermintaan/responData/dataKeuangan/dataRekening"):
        no_rekening = data_rek.findtext("infoRekening/noRekening", default="")
        mutasi_list = data_rek.findall("mutasiRekening")
        rekening_count[no_rekening] += len(mutasi_list)

    for data_rek in root.findall(".//responPermintaan/responData/dataKeuangan/dataRekening"):
        no_rekening = data_rek.findtext("infoRekening/noRekening", default="")
        dict_saldo_awal[no_rekening] = float(data_rek.findtext("infoRekening/saldoAwal", default="0"))
        daftar_rekening = {
            "no_rekening": no_rekening,
            "nama_rekening": data_rek.findtext("infoRekening/namaRekening", default=""),
            "tgl_pembukaan_rekening": data_rek.findtext("infoRekening/tglBukaRek", default=""),
            "tgl_penutupan_rekening": data_rek.findtext("infoRekening/tglTutupRek", default=""),
            "tgl_awal": data_rek.findtext("infoRekening/tglAwal", default=""),
            "tgl_akhir": data_rek.findtext("infoRekening/tglAkhir", default=""),
            "mata_uang": data_rek.findtext("infoRekening/mataUang", default=""),
            "status": data_rek.findtext("infoRekening/statusRekening", default=""),
            "saldo_awal": float(data_rek.findtext("infoRekening/saldoAwal", default="")),
            "saldo_akhir": float(data_rek.findtext("infoRekening/saldoAkhir", default="")),
            "jml_transaksi": rekening_count[no_rekening]
        }
        list_daftar_rekening.append(daftar_rekening)

    df_daftar_rekening = pd.DataFrame(list_daftar_rekening)

    # MUTASI PER REKENING
    data_per_rekening = {}
    for rekening in root.findall('.//responPermintaan/responData/dataKeuangan/dataRekening'):
        no_rekening = rekening.findtext('./mutasiRekening/noRekening')
        if no_rekening not in data_per_rekening:
            data_per_rekening[no_rekening] = []

        for mutasi in rekening.findall('./mutasiRekening'):
            debit_credit = mutasi.findtext('./kodeDebitCredit')
            try:
                nilai_transaksi = float(mutasi.findtext('./nilaiTransaksi', default='0'))
                debit = nilai_transaksi if debit_credit == 'dr' else 0.0
                credit = nilai_transaksi if debit_credit == 'cr' else 0.0
                saldo = dict_saldo_awal[no_rekening] - debit + credit
                dict_saldo_awal[no_rekening] = saldo
            except (ValueError, TypeError):
                nilai_transaksi = debit = credit = 0.0
                saldo = dict_saldo_awal[no_rekening]

            row = {
                'tanggal': mutasi.findtext('./tglTransaksi'),
                'no_rekening': no_rekening,
                'kd_jenis_transaksi': mutasi.findtext('./kdJnsTrans'),
                'kd_bank_lawan': mutasi.findtext('./kdBankLawan'),
                'no_rekening_lawan': mutasi.findtext('./noRekeningLawan'),
                'nama_rekening_lawan': mutasi.findtext('./namaRekeningLawan'),
                'debit_credit': debit_credit,
                'nilai_transaksi': nilai_transaksi,
                'debit': debit,
                'credit': credit,
                'saldo': saldo,
                'berita': mutasi.findtext('./berita')
            }
            data_per_rekening[no_rekening].append(row)

    # Write to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_info_respon.to_excel(writer, sheet_name='INFO RESPON', index=False)
        df_daftar_wp.to_excel(writer, sheet_name='DAFTAR WP', index=False)
        df_daftar_rekening.to_excel(writer, sheet_name='DAFTAR REKENING', index=False)
        for i, (no_rekening, records) in enumerate(data_per_rekening.items(), start=1):
            df_rek_sheet = pd.DataFrame(records)
            sheet_name = f"REK-{i}"
            df_rek_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output

# Streamlit UI
st.title("📄 XML to Excel Extractor")

uploaded_file = st.file_uploader("Upload XML File", type="xml")



if uploaded_file:
    st.success("File uploaded successfully!")

    # Extract base name without extension
    base_name = os.path.splitext(uploaded_file.name)[0]
    output_filename = f"{base_name}_output.xlsx"

    if st.button("Extract XML to excel"):
        with st.spinner("⏳ Extracting data... Please wait..."):
            excel_bytes = extract_xml_to_excel(uploaded_file.read())

        st.success("✅ Extraction complete!")
        st.download_button(
            label="📥 Download Excel File",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
