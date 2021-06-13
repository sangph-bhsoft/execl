import React, { useState } from "react";
import * as XLSX from "xlsx";
import Excel from 'exceljs';
import { saveAs } from 'file-saver';
import "./App.css";

const hl = {
  'Giỏi' : 'G',
  'Khá' : 'K',
  'Yếu' : 'Y',
  'TB' : 'TB',
  'Kém' : 'KEM'
};

const hk ={
  'Tốt' : 'T',
  'Khá' : 'K',
  'Trung bình' : 'TB',
  'Yếu' : 'Y'
}

const td = {
  'Học sinh Giỏi': 'Giỏi',
  'Học sinh Tiên tiến' : 'Tiên tiến'
}

function App() {
  const [studentF, setStudentF] = useState([]);
  const [studentE, setStudentE] = useState([]);


  const readFileExel = (file) => {
    return new Promise((resolve, reject) => {
      var reader = new FileReader();
      reader.onload = function (e) {
        var data = e.target.result;
        let readedData = XLSX.read(data, { type: "binary" });
        const wsname = readedData.SheetNames[0];
        const ws = readedData.Sheets[wsname];

        /* Convert array to json*/
        const dataParse = XLSX.utils.sheet_to_json(ws, { header: 1 });
        resolve(dataParse);
      };
      reader.onerror  = ()=>{
        reject(reader.error)
      }
      reader.readAsBinaryString(file);
    });
  };

  const handleChangeFileFrom = async (e) => {
    try {
      const file = e.target.files[0];
      const data = await readFileExel(file);
      const students= [];
      if(data && data.length){
        let newData = data.slice(5,data.length-2);
        
        newData.forEach(st => {
          students.push({
            hoten : st[1],
            toan : st[3].replace(',','.'),
            vatly : st[4].replace(',','.'),
            hoahoc : st[5].replace(',','.'),
            sinhhoc : st[6].replace(',','.'),
            tinhoc : st[7].replace(',','.'),
            nguvan : st[8].replace(',','.'),
            lichsu : st[9].replace(',','.'),
            dialy : st[10].replace(',','.'),
            ngoaingu : st[11].replace(',','.'),
            congnghe : st[12].replace(',','.'),
            theduc : st[13].replace(',','.'),
            qpan : st[14].replace(',','.'),
            gdcd : st[15].replace(',','.'),
            tb : st[16].replace(',','.'),
            hl : st[17],
            hk : st[18],
            td : st[19]
          })
        });
        setStudentF(students);
        
      }
      
    } catch (error) {
      window.alert('lỗi file');
      return;
    }
  };

  const handleChangeFileTo = async (e) => {
     try {
      const file = e.target.files[0];
      const data = await readFileExel(file);
      const students = [];
      if(data && data.length){
        let newData = data.slice(2,data.length);
        newData.forEach(st => {
          students.push({
            ml : st[1],
            mhs : st[2],
            hoten : st[3],
            ngaysinh : st[4],
          
          })
        });
        setStudentE(students);
        
      }
    } catch (error) {
      window.alert('lỗi file');
      return;
    }
  };

  const handleUpload=()=>{
    let fileName = studentE[0].ml;
    const students = [];
    studentE.forEach(st=>{
     
      let index = studentF.findIndex(s=>s.hoten === st.hoten);
      if(index !== -1){
        let student = studentF[index];
        students.push(
          {
            ...st,
            ...student,
          }
        )
      }else {
        console.log(st);
        
      }
    })
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.columns  = [
      {
        key : 'stt',
        width : 10
      },
      {
        key : 'ml',
        width : 15
      },
      {
        key : 'mhs',
        width : 20
      },
      {
        key : 'hoten',
        width : 25
      },
      {
        key : 'ngaysinh',
        width : 20
      },
      {
        key : 'toan',
        width : 10
      },
      {
        key : 'vatly',
        width : 10
      },
      {
        key : 'hoahoc',
        width : 10
      },
      {
        key : 'sinhhoc',
        width : 10
      },
      {
        key : 'tinhoc',
        width : 10
      },
      {
        key : 'nguvan',
        width : 10
      },
      {
        key : 'lichsu',
        width : 10
      },
      {
        key : 'dialy',
        width : 10
      },
      {
        key : 'ngoaingu',
        width : 10
      },
      {
        key : 'congnghe',
        width : 10
      },
      {
        key : 'qpan',
        width : 10
      },
      {
        key : 'theduc',
        width : 10
      },
      {
        key : 'ngoaingu2',
        width : 10
      },
      {
        key : 'nghept',
        width : 10
      },
      {
        key : 'gdcd',
        width : 10
      },
      ,
      {
        key : 'tb',
        width : 10
      },
      ,
      {
        key : 'hl',
        width : 10
      },
      {
        key : 'hk',
        width : 10
      },
      {
        key : 'td',
        width : 10
      }
    ]
    worksheet.mergeCells('A1:A2');
    worksheet.mergeCells('B1:B2');
    worksheet.mergeCells('C1:C2');
    worksheet.mergeCells('D1:D2');
    worksheet.mergeCells('E1:E2');
    worksheet.mergeCells('F1:F2');
    worksheet.mergeCells('G1:G2');
    worksheet.mergeCells('H1:H2');
    worksheet.mergeCells('I1:I2');
    worksheet.mergeCells('J1:J2');
    worksheet.mergeCells('K1:K2');
    worksheet.mergeCells('L1:L2');
    worksheet.mergeCells('M1:M2');
    worksheet.mergeCells('N1:N2');
    worksheet.mergeCells('O1:O2');
    worksheet.mergeCells('P1:P2');
    worksheet.mergeCells('Q1:Q2');
    worksheet.mergeCells('T1:T2');
    worksheet.mergeCells('U1:U2');

    const row1 = worksheet.getRow(1);

    row1.values = [
      'STT',
      'Mã lớp',
      'Mã học sinh',
      'Họ tên',
      'Ngày Sinh',
      'Toán',
      'Vật lý',
      'Hóa học',
      'Sinh học',
      'Tin học',
      'Ngữ văn',
      'Lịch sử',
      'Địa lý',
      'Ngoại ngữ',
      'Công nghệ',
      'GD QP AN',
      'Thể dục',
      '',
      '',
      'GDCD',
      'ĐTB các môn'
    ]

    worksheet.mergeCells('R1:S1');
    worksheet.getCell('R1').value="Tự chọn";
    worksheet.getCell('R2').value="Ngoại ngữ 2";
    worksheet.getCell('S2').value="Nghề phổ thông";
    worksheet.mergeCells('V1:X1');
    worksheet.getCell('V1').value="Kết quả xếp loại và DH thi đua";
    worksheet.getCell('V2').value="Học lực";
    worksheet.getCell('W2').value="Hạnh kiểm";
    worksheet.getCell('X2').value="Danh hiệu thi đua";

    
    

    
    students.map((st,stt)=>{
      worksheet.addRow({
        stt : stt+1,
        ...st,
        hl : hl[st.hl],
        hk : hk[st.hk],
        td : td[st.td]
      })
    })
    

    workbook.xlsx
          .writeBuffer()
          .then(buffer =>
            saveAs(
              new Blob([buffer]),
              `${fileName}.xlsx`
            )
          )
          .catch(err => console.log('Error writing excel export', err));
    
  }

  const handleUploadCN=()=>{
    let fileName = studentE[0].ml;
    const students = [];
    studentE.forEach(st=>{
      let index = studentF.findIndex(s=>s.hoten === st.hoten);
      if(index !== -1){
        let student = studentF[index];
        students.push(
          {
            ...st,
            ...student,
          }
        )
      }else {
        console.log(st);
        
      }
    })
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    worksheet.columns  = [
      {
        key : 'stt',
        width : 10
      },
      {
        key : 'ml',
        width : 15
      },
      {
        key : 'mhs',
        width : 20
      },
      {
        key : 'hoten',
        width : 25
      },
      {
        key : 'ngaysinh',
        width : 20
      },
      {
        key : 'toan',
        width : 10
      },
      {
        key : 'vatly',
        width : 10
      },
      {
        key : 'hoahoc',
        width : 10
      },
      {
        key : 'sinhhoc',
        width : 10
      },
      {
        key : 'tinhoc',
        width : 10
      },
      {
        key : 'nguvan',
        width : 10
      },
      {
        key : 'lichsu',
        width : 10
      },
      {
        key : 'dialy',
        width : 10
      },
      {
        key : 'ngoaingu',
        width : 10
      },
      {
        key : 'congnghe',
        width : 10
      },
      {
        key : 'qpan',
        width : 10
      },
      {
        key : 'theduc',
        width : 10
      },
      {
        key : 'ngoaingu2',
        width : 10
      },
      {
        key : 'nghept',
        width : 10
      },
      {
        key : 'gdcd',
        width : 10
      },
      ,
      {
        key : 'tb',
        width : 10
      },
      ,
      {
        key : 'hl',
        width : 10
      },
      {
        key : 'hk',
        width : 10
      },
      {
        key : 'td',
        width : 10
      },
      {
        key : 'ngaynghi',
        width : 10
      },
      {
        key : 'lenlop',
        width : 10
      },
      {
        key : 'kiemtralai',
        width : 10
      }
    ]
    worksheet.mergeCells('A1:A2');
    worksheet.mergeCells('B1:B2');
    worksheet.mergeCells('C1:C2');
    worksheet.mergeCells('D1:D2');
    worksheet.mergeCells('E1:E2');
    worksheet.mergeCells('F1:F2');
    worksheet.mergeCells('G1:G2');
    worksheet.mergeCells('H1:H2');
    worksheet.mergeCells('I1:I2');
    worksheet.mergeCells('J1:J2');
    worksheet.mergeCells('K1:K2');
    worksheet.mergeCells('L1:L2');
    worksheet.mergeCells('M1:M2');
    worksheet.mergeCells('N1:N2');
    worksheet.mergeCells('O1:O2');
    worksheet.mergeCells('P1:P2');
    worksheet.mergeCells('Q1:Q2');
    worksheet.mergeCells('T1:T2');
    worksheet.mergeCells('U1:U2');
    worksheet.mergeCells('Y1:Y2');
    worksheet.mergeCells('Z1:Z2');
    worksheet.mergeCells('AA1:AA2');

    const row1 = worksheet.getRow(1);

    row1.values = [
      'STT',
      'Mã lớp',
      'Mã học sinh',
      'Họ tên',
      'Ngày Sinh',
      'Toán',
      'Vật lý',
      'Hóa học',
      'Sinh học',
      'Tin học',
      'Ngữ văn',
      'Lịch sử',
      'Địa lý',
      'Ngoại ngữ',
      'Công nghệ',
      'GD QP AN',
      'Thể dục',
      '',
      '',
      'GDCD',
      'ĐTB các môn',
      '',
      '',
      '',
      'TS ngày nghỉ học cả năm',
      'Được lên lớp',
      'Kiểm tra lại, rèn luyện HK trong hè'
    ]

    worksheet.mergeCells('R1:S1');
    worksheet.getCell('R1').value="Tự chọn";
    worksheet.getCell('R2').value="Ngoại ngữ 2";
    worksheet.getCell('S2').value="Nghề phổ thông";
    worksheet.mergeCells('V1:X1');
    worksheet.getCell('V1').value="Kết quả xếp loại và DH thi đua";
    worksheet.getCell('V2').value="Học lực";
    worksheet.getCell('W2').value="Hạnh kiểm";
    worksheet.getCell('X2').value="Danh hiệu thi đua";

    
    

    
    students.map((st,stt)=>{
      worksheet.addRow({
        stt : stt+1,
        ...st,
        hl : hl[st.hl],
        hk : hk[st.hk],
        td : td[st.td],
        
      })
    })
    

    workbook.xlsx
          .writeBuffer()
          .then(buffer =>
            saveAs(
              new Blob([buffer]),
              `${fileName}_CN_.xlsx`
            )
          )
          .catch(err => console.log('Error writing excel export', err));
    
  }

  return (
    
    <div className="App">
     <div className="container">
     <div className="file-wrapper">
        <label htmlFor="">
          File dữ liệu:
        </label> 
        <input type="file" onChange={handleChangeFileFrom} />
      </div>
      <div className="file-wrapper">
        <label htmlFor="">
           File Mẫu:
        </label> 
        <input type="file" onChange={handleChangeFileTo} />
      </div>
      
      <div className="btn">
        <button disabled={!studentE.length || !studentF.length} onClick={handleUpload}>Download File HK2</button>
      <button disabled={!studentE.length || !studentF.length} onClick={handleUploadCN}>Download File CN</button>
      </div>
     </div>
     
    </div>
  );
}

export default App;
