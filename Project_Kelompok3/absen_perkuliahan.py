from tkinter import *
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from tkinter import font as tkfont

class AplikasiAbsensi:
    def __init__(self, root):
        self.root = root
        self.root.title("Absensi Perkuliahan")
        self.root.geometry("800x600")

        self.workbook = Workbook()
        self.sheet = self.workbook.active

        self.num = 0
        self.border = Border(left=Side(border_style='thin', color='00000000'),
                             right=Side(border_style='thin', color='00000000'),
                             top=Side(border_style='thin', color='00000000'),
                             bottom=Side(border_style='thin', color='00000000'))
        self.alignment = Alignment(horizontal='center', vertical='center')

        self.initialize_gui()

    def initialize_gui(self):
        styling = tkfont.Font(family='Helvetica', weight='bold', size=15)
        styling2 = tkfont.Font(family='Helvetica', size=9)

        font = Font(bold=True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        gui_width = int(0.8 * screen_width)
        gui_height = int(0.8 * screen_height)

        self.root.geometry(f"{gui_width}x{gui_height}")

        canvas = Canvas(self.root, bg='grey')
        canvas.pack(fill=BOTH, expand=YES)

        self.sheet['A1'] = "Mata Kuliah\t:"
        A1 = self.sheet['A1']
        A1.font = font
        self.sheet['A2'] = "Tanggal Perkuliahan\t:"
        A2 = self.sheet['A2']
        A2.font = font

        self.sheet['A3'] = "No"
        A3 = self.sheet['A3']
        A3.font = font
        A3.border = self.border
        A3.alignment = self.alignment

        self.sheet['B3'] = "Nama"
        B3 = self.sheet['B3']
        B3.font = font
        B3.border = self.border
        B3.alignment = self.alignment

        self.sheet['C3'] = "NIM"
        C3 = self.sheet['C3']
        C3.font = font
        C3.border = self.border
        C3.alignment = self.alignment

        self.sheet['D3'] = "Jurusan"
        D3 = self.sheet['D3']
        D3.font = font
        D3.border = self.border
        D3.alignment = self.alignment

        frameJudul = Frame(self.root, bg='white')
        frameJudul.place(rely=0.025, relx=0.5, relheight=0.1, relwidth=0.8, anchor='n')
        judul = Label(frameJudul, bg='white', text='Absensi Perkuliahan', font=styling)
        judul.place(relheight=1, relwidth=1)

        frameMatkul = Frame(self.root, bg='white')
        frameMatkul.place(rely=0.2, relx=0.5, relheight=0.06, relwidth=0.8, anchor='n')
        matkulinfo = Label(frameMatkul, bg='white', text='Mata kuliah', font=styling2)
        matkulinfo.place(relwidth=0.4, relheight=1)
        self.matkulEntry = Entry(frameMatkul)
        self.matkulEntry.place(relx=0.4, relheight=1, relwidth=0.6)

        frameTanggal = Frame(self.root, bg='white')
        frameTanggal.place(rely=0.27, relx=0.5, relheight=0.06, relwidth=0.8, anchor='n')
        tanggalinfo = Label(frameTanggal, bg='white', text='Tanggal Perkuliahan', font=styling2)
        tanggalinfo.place(relwidth=0.4, relheight=1)
        self.tanggalEntry = Entry(frameTanggal)
        self.tanggalEntry.place(relx=0.4, relheight=1, relwidth=0.6)

        frameNama = Frame(self.root, bg='white')
        frameNama.place(rely=0.34, relx=0.5, relheight=0.06, relwidth=0.8, anchor='n')
        namainfo = Label(frameNama, bg='white', text='Nama', font=styling2)
        namainfo.place(relwidth=0.4, relheight=1)
        self.namaEntry = Entry(frameNama)
        self.namaEntry.place(relx=0.4, relheight=1, relwidth=0.6)

        frameNIM = Frame(self.root, bg='white')
        frameNIM.place(rely=0.41, relx=0.5, relheight=0.06, relwidth=0.8, anchor='n')
        NIMinfo = Label(frameNIM, bg='white', text='NIM', font=styling2)
        NIMinfo.place(relwidth=0.4, relheight=1)
        self.NIMEntry = Entry(frameNIM)
        self.NIMEntry.place(relx=0.4, relheight=1, relwidth=0.6)

        frameJurusan = Frame(self.root, bg='white')
        frameJurusan.place(rely=0.48, relx=0.5, relheight=0.06, relwidth=0.8, anchor='n')
        jurusaninfo = Label(frameJurusan, bg='white', text='Jurusan', font=styling2)
        jurusaninfo.place(relwidth=0.4, relheight=1)
        self.jurusanEntry = Entry(frameJurusan)
        self.jurusanEntry.place(relx=0.4, relheight=1, relwidth=0.6)

        self.informasi = Label(self.root, bg='white', font=styling2, text='Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.')
        self.informasi.place(rely=0.56, relx=0.5, relheight=0.1, relwidth=0.8, anchor='n')

        frameButton = Frame(self.root, bg='white')
        frameButton.place(rely=0.675, relx=0.5, relheight=0.3, relwidth=0.3, anchor='n')
        insert = Button(frameButton, text='Insert', command=self.InsertData)
        insert.place(rely=0, relx=0.5, relheight=0.25, relwidth=1, anchor='n')
        save = Button(frameButton, text='Save', command=self.SaveData)
        save.place(rely=0.25, relx=0.5, relheight=0.25, relwidth=1, anchor='n')
        createNewData = Button(frameButton, text='Create New', command=self.CreateNewData)
        createNewData.place(rely=0.5, relx=0.5, relheight=0.25, relwidth=1, anchor='n')
        Exit = Button(frameButton, text='Exit', command=self.root.quit)
        Exit.place(rely=0.75, relx=0.5, relheight=0.25, relwidth=1, anchor='n')

    def InsertData(self):
        self.num = self.num + 1
        sheetnum = self.num + 3

        self.sheet['A' + str(sheetnum)] = self.num
        DataNo = self.sheet['A' + str(sheetnum)]
        DataNo.border = self.border
        DataNo.alignment = self.alignment

        self.sheet['B' + str(sheetnum)] = self.namaEntry.get()
        DataNama = self.sheet['B' + str(sheetnum)]
        DataNama.border = self.border
        DataNama.alignment = self.alignment

        self.sheet['C' + str(sheetnum)] = self.NIMEntry.get()
        DataNIM = self.sheet['C' + str(sheetnum)]
        DataNIM.border = self.border
        DataNIM.alignment = self.alignment

        self.sheet['D' + str(sheetnum)] = self.jurusanEntry.get()
        DataJurusan = self.sheet['D' + str(sheetnum)]
        DataJurusan.border = self.border
        DataJurusan.alignment = self.alignment

        self.sheet['B1'] = self.matkulEntry.get()
        self.sheet['B2'] = self.tanggalEntry.get()

        self.namaEntry.delete(0, END)
        self.NIMEntry.delete(0, END)
        self.jurusanEntry.delete(0, END)

    def SaveData(self):
        self.workbook.save(filename=str(self.matkulEntry.get()) + "_" + str(self.tanggalEntry.get()) + ".xlsx")
        self.informasi['text'] = "Data absen telah di save!\nNama file: " + str(self.matkulEntry.get()) + "_" + str(
            self.tanggalEntry.get()) + ".xlsx"

    def CreateNewData(self):
        self.informasi['text'] = 'Klik Insert untuk semua mahasiswa, kemudian klik Save jika semua telah diabsen.'
        self.namaEntry.delete(0, END)
        self.NIMEntry.delete(0, END)
        self.jurusanEntry.delete(0, END)
        self.matkulEntry.delete(0, END)
        self.tanggalEntry.delete(0, END)
        self.num = 0

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    root = Tk()
    aplikasi = AplikasiAbsensi(root)
    aplikasi.run()
