'use client';

import { useState, useEffect, useRef } from 'react';
import { useDropzone } from 'react-dropzone';
import { read, utils } from 'xlsx';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

export default function Home() {
  const [excelFile, setExcelFile] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);
  const [participants, setParticipants] = useState([]);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');
  const [progress, setProgress] = useState(0);
  const [timeRemaining, setTimeRemaining] = useState(null);
  const [startTime, setStartTime] = useState(null);
  const [activeStep, setActiveStep] = useState(1);
  const [showVBAModal, setShowVBAModal] = useState(false);
  const [sourceFolder, setSourceFolder] = useState("");
  const [targetFolder, setTargetFolder] = useState("");

  const codeRef = useRef(null);

  // VBA Code Template
  // VBA Code Template
const getVBACode = () => {
  // ตรวจสอบและเพิ่ม \ ต่อท้าย path ถ้าไม่มี
  let srcFolder = sourceFolder;
  let tgtFolder = targetFolder;
  
  if (!srcFolder.endsWith('\\')) {
    srcFolder += '\\';
  }
  
  if (!tgtFolder.endsWith('\\')) {
    tgtFolder += '\\';
  }
  
  return `Sub BatchConvertDocToPDF()
    Dim doc As Document
    Dim sourceFolder As String
    Dim targetFolder As String
    Dim file As String
    Dim docName As String
    Dim pdfName As String
    
    ' กำหนดโฟลเดอร์ต้นทางและโฟลเดอร์ปลายทาง
    sourceFolder = "${srcFolder}"
    targetFolder = "${tgtFolder}"

    ' รับไฟล์ .docx ทั้งหมดจากโฟลเดอร์ต้นทาง
    file = Dir(sourceFolder & "*.docx")

    ' ทำการวนลูปเพื่อแปลงไฟล์ทั้งหมด
    Do While file <> ""
        ' เปิดไฟล์เอกสาร
        Set doc = Documents.Open(sourceFolder & file)
        
        ' กำหนดชื่อไฟล์ PDF
        docName = Left(file, Len(file) - 5)
        pdfName = targetFolder & docName & ".pdf"
        
        ' บันทึกเป็น PDF
        doc.ExportAsFixedFormat OutputFileName:=pdfName, ExportFormat:=wdExportFormatPDF
        
        ' ปิดเอกสาร
        doc.Close False
        
        ' อ่านไฟล์ถัดไป
        file = Dir
    Loop
    
    ' แสดงข้อความเมื่อเสร็จสิ้น
    MsgBox "การแปลงไฟล์ทั้งหมดเสร็จสิ้น!"
End Sub`;
};

  // คัดลอกโค้ด VBA
  const copyVBACode = () => {
    const code = getVBACode();
    navigator.clipboard.writeText(code)
      .then(() => {
        // สร้าง feedback ชั่วคราว
        const copyFeedback = document.createElement('div');
        copyFeedback.textContent = 'คัดลอกโค้ดแล้ว!';
        copyFeedback.className = 'fixed top-24 left-1/2 transform -translate-x-1/2 bg-green-600 text-white px-4 py-2 rounded-lg shadow-lg z-50 animate-fadeIn';
        document.body.appendChild(copyFeedback);

        // ลบหลังจาก 2 วินาที
        setTimeout(() => {
          copyFeedback.classList.add('opacity-0', 'transition-opacity', 'duration-300');
          setTimeout(() => document.body.removeChild(copyFeedback), 300);
        }, 2000);
      })
      .catch(err => {
        console.error('ไม่สามารถคัดลอกโค้ดได้:', err);
      });
  };

  // คำนวณเวลาที่เหลือ
  useEffect(() => {
    let timer;
    if (isLoading && progress > 0 && progress < 100) {
      timer = setInterval(() => {
        const elapsedTime = (Date.now() - startTime) / 1000; // เวลาที่ผ่านไปเป็นวินาที
        const totalEstimatedTime = (elapsedTime / progress) * 100; // เวลาทั้งหมดที่คาดการณ์
        const remaining = Math.max(0, totalEstimatedTime - elapsedTime); // เวลาที่เหลือ

        setTimeRemaining(Math.round(remaining));
      }, 1000);
    }

    return () => clearInterval(timer);
  }, [isLoading, progress, startTime]);

  // อัพเดท active step ตามสถานะปัจจุบัน
  useEffect(() => {
    if (excelFile && !templateFile) {
      setActiveStep(2);
    } else if (excelFile && templateFile) {
      setActiveStep(3);
    }
  }, [excelFile, templateFile]);

  // ไฮไลท์โค้ด VBA
  useEffect(() => {
    if (showVBAModal && codeRef.current) {
      // หากมีไลบรารีไฮไลท์โค้ด สามารถเรียกใช้ที่นี่
      // ตัวอย่างเช่น: Prism.highlightElement(codeRef.current);
    }
  }, [showVBAModal]);

  // อ่านไฟล์ Excel
  const readExcelFile = async (file) => {
    try {
      const data = await file.arrayBuffer();
      const workbook = read(data);
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = utils.sheet_to_json(worksheet);

      return jsonData;
    } catch (error) {
      console.error('Error reading Excel file:', error);
      throw new Error('ไม่สามารถอ่านไฟล์ Excel ได้');
    }
  };

  // สร้าง Certificate จาก Word template
  const generateCertificate = async (templateFile, data) => {
    try {
      // อ่าน template file
      const templateArrayBuffer = await templateFile.arrayBuffer();
      const zip = new PizZip(templateArrayBuffer);

      // สร้าง docxtemplater instance
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // เซ็ตข้อมูลให้กับ template
      doc.setData(data);

      // สร้าง document
      doc.render();

      // ดึงไฟล์ output เป็น blob
      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });

      return out;
    } catch (error) {
      console.error('Error generating certificate:', error);
      throw new Error('ไม่สามารถสร้าง certificate ได้: ' + error.message);
    }
  };

  // สร้าง Certificates หลายรายการ
  const generateMultipleCertificates = async (templateFile, dataList) => {
    try {
      // ถ้ามีข้อมูลหลายรายการให้สร้าง zip file
      if (dataList.length > 1) {
        const JSZip = (await import('jszip')).default;
        const zip = new JSZip();

        // เริ่มต้นการจับเวลา
        setStartTime(Date.now());
        setProgress(0);

        for (const [index, data] of dataList.entries()) {
          const certificate = await generateCertificate(templateFile, data);

          // ดึงค่าคอลัมน์แรกจากข้อมูล
          const firstColumnKey = Object.keys(data)[0];
          const firstColumnValue = data[firstColumnKey];

          // สร้างชื่อไฟล์จากเลขที่ + ค่าในคอลัมน์แรก
          const fileName = `${index + 1}.${firstColumnValue}.docx`;

          zip.file(fileName, certificate);

          // อัพเดทความคืบหน้า
          const currentProgress = Math.round(((index + 1) / dataList.length) * 100);
          setProgress(currentProgress);

          // หน่วงเวลาสั้นๆ เพื่อให้ UI อัพเดท
          await new Promise(resolve => setTimeout(resolve, 10));
        }

        // สร้างไฟล์ zip
        setProgress(95); // กำลังสร้างไฟล์ zip
        const content = await zip.generateAsync({
          type: 'blob',
          streamFiles: true
        }, (metadata) => {
          // อัพเดทความคืบหน้าในการสร้างไฟล์ zip
          setProgress(95 + Math.round(metadata.percent * 0.05));
        });

        setProgress(100);
        saveAs(content, 'certificates.zip');
        return 'สร้าง Certificates ทั้งหมด ' + dataList.length + ' รายการเรียบร้อยแล้ว';
      } else if (dataList.length === 1) {
        // ถ้ามีข้อมูลเพียงรายการเดียว
        setProgress(30);
        const certificate = await generateCertificate(templateFile, dataList[0]);

        // ดึงค่าคอลัมน์แรกจากข้อมูล
        const firstColumnKey = Object.keys(dataList[0])[0];
        const firstColumnValue = dataList[0][firstColumnKey];

        // สร้างชื่อไฟล์จากเลขที่ + ค่าในคอลัมน์แรก
        const fileName = `1.${firstColumnValue}.docx`;

        setProgress(90);
        saveAs(certificate, fileName);
        setProgress(100);
        return 'สร้าง Certificate เรียบร้อยแล้ว';
      } else {
        throw new Error('ไม่มีข้อมูลสำหรับสร้าง Certificate');
      }
    } catch (error) {
      console.error('Error generating certificates:', error);
      throw new Error('ไม่สามารถสร้าง certificates ได้: ' + error.message);
    }
  };

  // จัดการอัพโหลดไฟล์ Excel
  const handleExcelUpload = async (file) => {
    try {
      setExcelFile(file);
      const data = await readExcelFile(file);
      setParticipants(data);
      setError('');
      setSuccess('อัพโหลดไฟล์ Excel เรียบร้อยแล้ว: ' + file.name);
    } catch (err) {
      setError('ไม่สามารถอ่านไฟล์ Excel ได้: ' + err.message);
      setSuccess('');
    }
  };

  // จัดการอัพโหลดไฟล์ Word Template
  const handleTemplateUpload = (file) => {
    setTemplateFile(file);
    setSuccess((prev) => prev ? `${prev}\nอัพโหลดไฟล์ Template เรียบร้อยแล้ว: ${file.name}` : `อัพโหลดไฟล์ Template เรียบร้อยแล้ว: ${file.name}`);
  };

  // จัดการการสร้าง Certificates
  const handleGenerateCertificates = async () => {
    if (!excelFile) {
      setError('กรุณาอัพโหลดไฟล์ Excel ก่อน');
      return;
    }

    if (!templateFile) {
      setError('กรุณาอัพโหลดไฟล์ Template Word ก่อน');
      return;
    }

    try {
      setIsLoading(true);
      setError('');
      setProgress(0);
      setTimeRemaining(null);

      const result = await generateMultipleCertificates(templateFile, participants);

      setSuccess(result);
      setIsLoading(false);
    } catch (err) {
      setError('เกิดข้อผิดพลาดในการสร้าง Certificate: ' + err.message);
      setSuccess('');
      setIsLoading(false);
      setProgress(0);
      setTimeRemaining(null);
    }
  };

  // แปลงวินาทีเป็นรูปแบบเวลา
  const formatTime = (seconds) => {
    if (seconds === null) return '';
    if (seconds < 60) return `${seconds} วินาที`;
    return `${Math.floor(seconds / 60)} นาที ${seconds % 60} วินาที`;
  };

  // Component สำหรับ Drop Zone อัพโหลดไฟล์
  const FileUploader = ({ onFileUpload, accept, label, icon, isActive }) => {
    const { getRootProps, getInputProps, isDragActive } = useDropzone({
      onDrop: (acceptedFiles) => {
        if (acceptedFiles?.[0]) {
          onFileUpload(acceptedFiles[0]);
        }
      },
      accept,
      maxFiles: 1,
    });

    return (
      <div
        {...getRootProps()}
        className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-all duration-300 
          ${isDragActive ? 'border-blue-500 bg-blue-100' :
            isActive ? 'border-blue-400 bg-blue-50 shadow-md hover:shadow-lg hover:border-blue-500 hover:bg-blue-100' :
              'border-gray-300 bg-gray-50 hover:bg-gray-100 hover:border-gray-400'}`}
      >
        <input {...getInputProps()} />
        <div className="flex flex-col items-center">
          <div className={`transition-transform duration-300 ${isDragActive ? 'scale-110' : isActive ? 'scale-105' : ''}`}>
            {icon}
          </div>
          {
            isDragActive ?
              <p className="text-blue-600 font-medium mt-3 animate-pulse">วางไฟล์ที่นี่...</p> :
              <p className={`mt-3 ${isActive ? 'text-blue-700 font-medium' : 'text-gray-600'}`}>{label}</p>
          }
          <p className="text-xs text-gray-500 mt-2">คลิกหรือลากไฟล์มาวางที่นี่</p>
        </div>
      </div>
    );
  };

  // Component สำหรับ Step Indicator
  const StepIndicator = ({ steps, activeStep }) => {
    return (
      <div className="flex items-center justify-center w-full mb-12 flex-wrap">
        {steps.map((step, index) => (
          <div key={index} className="flex items-center mb-4">
            {/* Step circle */}
            <div
              className={`flex items-center justify-center w-10 h-10 rounded-full border-2 transition-all duration-300
                ${activeStep > index ? 'bg-green-100 border-green-500 text-green-700' :
                  activeStep === index ? 'bg-blue-100 border-blue-500 text-blue-700 shadow-md' :
                    'bg-gray-100 border-gray-300 text-gray-500'}`}
            >
              {activeStep > index ? (
                <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
                </svg>
              ) : (
                <span className="text-sm font-semibold">{index + 1}</span>
              )}
            </div>

            {/* Step label */}
            <div className="ml-3 mr-10">
              <p className={`text-sm font-medium mb-0.5 ${activeStep === index ? 'text-blue-700' :
                  activeStep > index ? 'text-green-700' : 'text-gray-500'
                }`}>
                {step.title}
              </p>
              <p className="text-xs text-gray-500 max-w-[140px]">{step.description}</p>
            </div>

            {/* Connector line */}
            {index < steps.length - 1 && (
              <div className={`hidden md:block flex-1 h-0.5 w-10 ${activeStep > index ? 'bg-green-500' : 'bg-gray-300'
                }`}></div>
            )}
          </div>
        ))}
      </div>
    );
  };

  // ข้อมูลสำหรับ Step Indicator
  const steps = [
    { title: 'อัพโหลด Excel', description: 'ไฟล์ข้อมูลรายชื่อ' },
    { title: 'อัพโหลด Template', description: 'ไฟล์ Word Template' },
    { title: 'สร้าง Certificate', description: 'ประมวลผลและดาวน์โหลด' }
  ];

  // Modal สำหรับโค้ด VBA
  const VBAModal = () => {
    if (!showVBAModal) return null;

    return (
      <div className="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4 animate-fadeIn">
        <div className="bg-white rounded-xl shadow-2xl w-full max-w-3xl max-h-[90vh] flex flex-col">
          <div className="p-6 bg-gradient-to-r from-blue-600 to-blue-500 rounded-t-xl flex justify-between items-center">
            <h3 className="text-xl font-semibold text-white">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 inline-block mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4" />
              </svg>
              โค้ด VBA สำหรับแปลงไฟล์ Word เป็น PDF
            </h3>
            <button
              className="text-white hover:bg-blue-700 rounded-full p-1 transition-colors"
              onClick={() => setShowVBAModal(false)}
            >
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
              </svg>
            </button>
          </div>

          <div className="p-6 flex-grow overflow-y-auto">
            <div className="mb-6">
              <h4 className="text-gray-700 font-semibold mb-2">กำหนดโฟลเดอร์:</h4>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-gray-600 text-sm mb-1">โฟลเดอร์ต้นทาง (ไฟล์ Word)</label>
                  <input
                    type="text"
                    value={sourceFolder}
                    onChange={(e) => setSourceFolder(e.target.value)}
                    className="w-full p-2 border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                    placeholder="ระบุโฟลเดอร์ต้นทาง"
                  />
                </div>
                <div>
                  <label className="block text-gray-600 text-sm mb-1">โฟลเดอร์ปลายทาง (ไฟล์ PDF)</label>
                  <input
                    type="text"
                    value={targetFolder}
                    onChange={(e) => setTargetFolder(e.target.value)}
                    className="w-full p-2 border border-gray-300 rounded focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm"
                    placeholder="ระบุโฟลเดอร์ปลายทาง"
                  />
                </div>
              </div>
            </div>

            <div className="mb-4">
              <div className="flex justify-between items-center mb-2">
                <h4 className="text-gray-700 font-semibold">โค้ด VBA:</h4>
                <button
                  className="bg-blue-600 hover:bg-blue-700 text-white text-sm px-3 py-1 rounded flex items-center transition-colors"
                  onClick={copyVBACode}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3" />
                  </svg>
                  คัดลอกโค้ด
                </button>
              </div>
              <div className="bg-gray-900 text-gray-100 p-4 rounded-md overflow-x-auto">
                <pre className="text-sm font-mono"><code ref={codeRef}>{getVBACode()}</code></pre>
              </div>
            </div>

            <div className="bg-blue-50 border-l-4 border-blue-500 p-4 rounded-md">
              <h4 className="text-blue-700 font-semibold mb-2 flex items-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-1" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                วิธีใช้งาน
              </h4>
              <ol className="list-decimal ml-5 text-sm text-gray-700 space-y-1">
                <li>เปิด Microsoft Word</li>
                <li>กด Alt + F11 เพื่อเปิด Visual Basic for Applications (VBA) editor</li>
                <li>ไปที่ Insert - Module เพื่อสร้างโมดูลใหม่</li>
                <li>วางโค้ดด้านบนลงในโมดูล</li>
                <li>กด F5 หรือคลิกปุ่ม Run เพื่อเริ่มการทำงาน</li>
                <li>รอจนกว่าโปรแกรมจะแสดงข้อความว่าการแปลงไฟล์เสร็จสิ้น</li>
              </ol>
            </div>
          </div>

          <div className="p-4 bg-gray-50 rounded-b-xl border-t">
            <div className="flex justify-end space-x-2">
              <button
                onClick={() => setShowVBAModal(false)}
                className="px-4 py-2 border border-gray-300 rounded-md text-gray-700 hover:bg-gray-100 transition-colors"
              >
                ปิด
              </button>
              <button
                onClick={copyVBACode}
                className="px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors"
              >
                คัดลอกโค้ด
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-blue-50 via-white to-blue-50">
      <main className="container mx-auto py-10 px-4 max-w-6xl">
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold text-blue-800 mb-3">ระบบสร้าง Certificate</h1>
          <p className="text-gray-600 max-w-2xl mx-auto">สร้าง Certificate จากไฟล์ Excel และ Template Word ได้อย่างรวดเร็วและง่ายดาย</p>
        </div>

        <StepIndicator steps={steps} activeStep={activeStep} />

        <div className="bg-white rounded-xl shadow-lg overflow-hidden mb-8">
          <div className="bg-gradient-to-r from-blue-600 to-blue-500 px-6 py-4">
            <h2 className="text-xl font-semibold text-white">การอัพโหลดไฟล์</h2>
          </div>

          <div className="p-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className={`transition-all duration-300 ${activeStep === 1 ? 'scale-100 opacity-100' : 'scale-95 opacity-80'}`}>
                <h3 className="text-lg font-medium mb-4 flex items-center text-blue-800">
                  <div className="flex items-center justify-center w-7 h-7 rounded-full bg-blue-100 text-blue-800 mr-3">1</div>
                  อัพโหลดไฟล์รายชื่อ (Excel)
                </h3>
                <FileUploader
                  onFileUpload={handleExcelUpload}
                  accept={{
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    'application/vnd.ms-excel': ['.xls'],
                  }}
                  label="อัพโหลดไฟล์ Excel"
                  isActive={activeStep === 1}
                  icon={
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-16 w-16 text-green-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                    </svg>
                  }
                />
                {excelFile && (
                  <div className="mt-4 bg-green-50 p-4 rounded-lg border border-green-200 transform transition-all duration-300 hover:shadow-md">
                    <div className="flex items-center">
                      <div className="flex-shrink-0 h-10 w-10 rounded-full bg-green-100 flex items-center justify-center">
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-green-600" viewBox="0 0 20 20" fill="currentColor">
                          <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                        </svg>
                      </div>
                      <div className="ml-4">
                        <p className="text-sm font-medium text-green-800">อัพโหลดสำเร็จ</p>
                        <p className="text-sm text-green-700">{excelFile.name}</p>
                        <p className="text-xs text-green-600 mt-1">จำนวนข้อมูล: {participants.length} รายการ</p>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              <div className={`transition-all duration-300 ${activeStep === 2 ? 'scale-100 opacity-100' : 'scale-95 opacity-80'}`}>
                <h3 className="text-lg font-medium mb-4 flex items-center text-blue-800">
                  <div className="flex items-center justify-center w-7 h-7 rounded-full bg-blue-100 text-blue-800 mr-3">2</div>
                  อัพโหลดไฟล์ Template (Word)
                </h3>
                <FileUploader
                  onFileUpload={handleTemplateUpload}
                  accept={{
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                  }}
                  label="อัพโหลดไฟล์ Word Template"
                  isActive={activeStep === 2}
                  icon={
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-16 w-16 text-blue-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5} d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
                    </svg>
                  }
                />
                {templateFile && (
                  <div className="mt-4 bg-blue-50 p-4 rounded-lg border border-blue-200 transform transition-all duration-300 hover:shadow-md">
                    <div className="flex items-center">
                      <div className="flex-shrink-0 h-10 w-10 rounded-full bg-blue-100 flex items-center justify-center">
                        <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-blue-600" viewBox="0 0 20 20" fill="currentColor">
                          <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                        </svg>
                      </div>
                      <div className="ml-4">
                        <p className="text-sm font-medium text-blue-800">อัพโหลดสำเร็จ</p>
                        <p className="text-sm text-blue-700">{templateFile.name}</p>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>

        {participants.length > 0 && (
          <div className="bg-white rounded-xl shadow-lg overflow-hidden mb-8 transform transition-all duration-500">
            <div className="bg-gradient-to-r from-indigo-600 to-indigo-500 px-6 py-4 flex justify-between items-center">
              <h2 className="text-xl font-semibold text-white flex items-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
                  <path d="M3 4a1 1 0 011-1h12a1 1 0 011 1v2a1 1 0 01-1 1H4a1 1 0 01-1-1V4zM3 10a1 1 0 011-1h6a1 1 0 011 1v6a1 1 0 01-1 1H4a1 1 0 01-1-1v-6zM14 9a1 1 0 00-1 1v6a1 1 0 001 1h2a1 1 0 001-1v-6a1 1 0 00-1-1h-2z" />
                </svg>
                ข้อมูลที่อ่านได้จาก Excel
              </h2>
              <span className="text-white bg-indigo-800 px-3 py-1 rounded-full text-xs font-medium">
                {participants.length} รายการ
              </span>
            </div>
            <div className="p-6">
              <div className="bg-gray-50 rounded-lg border border-gray-200 max-h-80 overflow-auto shadow-inner">
                <table className="min-w-full divide-y divide-gray-200">
                  <thead className="bg-gray-100 sticky top-0 shadow-sm z-10">
                    <tr>
                      <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider w-16">ลำดับ</th>
                      {Object.keys(participants[0]).map((key) => (
                        <th key={key} className="px-4 py-3 text-left text-xs font-medium text-gray-600">
                          {key}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {participants.slice(0, 5).map((item, index) => (
                      <tr key={index} className="hover:bg-blue-50 transition-colors">
                        <td className="px-4 py-3 whitespace-nowrap text-sm font-medium text-gray-700">
                          {index + 1}
                        </td>
                        {Object.values(item).map((value, i) => (
                          <td key={i} className="px-4 py-3 whitespace-nowrap text-sm text-gray-700">
                            {String(value)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {participants.length > 5 && (
                  <div className="text-sm text-gray-500 text-center bg-gray-100 py-2 border-t border-gray-200">
                    แสดง 5 รายการแรกจากทั้งหมด {participants.length} รายการ
                  </div>
                )}
              </div>
              <div className="mt-5 bg-yellow-50 p-4 rounded-lg border border-yellow-200">
                <h4 className="font-medium text-yellow-800 flex items-center">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2 text-yellow-600" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2h-1V9z" clipRule="evenodd" />
                  </svg>
                  คำแนะนำสำคัญ
                </h4>
                <ul className="mt-3 space-y-2 text-sm text-yellow-700">
                  <li className="flex items-start">
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-1.5 mt-0.5 text-yellow-600" viewBox="0 0 20 20" fill="currentColor">
                      <path fillRule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clipRule="evenodd" />
                    </svg>
                    <span>ชื่อคอลัมน์ในไฟล์ Excel ต้องตรงกับ placeholders ในไฟล์ Word Template โดยใช้รูปแบบ <code className="bg-yellow-100 px-1.5 py-0.5 rounded text-yellow-800 font-mono">{`{ชื่อคอลัมน์}`}</code></span>
                  </li>
                  <li className="flex items-start">
                    <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-1.5 mt-0.5 text-yellow-600" viewBox="0 0 20 20" fill="currentColor">
                      <path fillRule="evenodd" d="M7.293 14.707a1 1 0 010-1.414L10.586 10 7.293 6.707a1 1 0 011.414-1.414l4 4a1 1 0 010 1.414l-4 4a1 1 0 01-1.414 0z" clipRule="evenodd" />
                    </svg>
                    <span>ไฟล์ที่สร้างจะตั้งชื่อเป็น <code className="bg-yellow-100 px-1.5 py-0.5 rounded text-yellow-800 font-mono">[ลำดับ].[ค่าในคอลัมน์แรก].docx</code></span>
                  </li>
                </ul>
              </div>
            </div>
          </div>
        )}

        {isLoading && progress > 0 && (
          <div className="bg-white rounded-xl shadow-lg overflow-hidden mb-8 animate-fadeIn">
            <div className="bg-gradient-to-r from-blue-600 to-blue-500 px-6 py-4 flex justify-between items-center">
              <h2 className="text-xl font-semibold text-white flex items-center">
                <svg className="animate-spin h-5 w-5 mr-3 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                กำลังสร้าง Certificate
              </h2>
              <div className="text-white font-medium bg-blue-800 px-3 py-1 rounded-full text-sm">
                {progress}%
              </div>
            </div>
            <div className="p-6">
              <div className="relative pt-1">
                <div className="overflow-hidden h-4 mb-4 text-xs flex rounded-full bg-blue-100">
                  <div
                    style={{ width: `${progress}%` }}
                    className="shadow-lg flex flex-col text-center whitespace-nowrap text-white justify-center bg-gradient-to-r from-blue-500 to-blue-600 transition-all duration-300 ease-out"
                  >
                    <div className="absolute top-0 left-0 right-0 h-4 flex items-center justify-center text-xs font-semibold text-blue-100">
                      {progress < 10 ? '' : `${progress}%`}
                    </div>
                  </div>
                </div>
              </div>

              <div className="flex flex-col md:flex-row justify-between items-center text-sm text-gray-600 bg-blue-50 p-4 rounded-lg border border-blue-200">
                <div className="flex items-center mb-2 md:mb-0">
                  <div className="h-8 w-8 rounded-full bg-blue-100 flex items-center justify-center mr-3">
                    <svg className="animate-spin h-4 w-4 text-blue-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                    </svg>
                  </div>
                  <span>
                    กำลังประมวลผล {participants.length} รายการ
                  </span>
                </div>
                <div className="flex items-center">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-1.5 text-blue-500" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm1-12a1 1 0 10-2 0v4a1 1 0 00.293.707l2.828 2.829a1 1 0 101.415-1.415L11 9.586V6z" clipRule="evenodd" />
                  </svg>
                  {timeRemaining !== null ? (
                    <span>เวลาที่เหลือโดยประมาณ: <strong>{formatTime(timeRemaining)}</strong></span>
                  ) : (
                    <span>กำลังคำนวณเวลาที่เหลือ...</span>
                  )}
                </div>
              </div>

              {progress > 0 && progress < 100 && (
                <div className="mt-4 text-xs text-gray-500 flex items-center">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-4 w-4 mr-1.5 text-gray-400" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2h-1V9z" clipRule="evenodd" />
                  </svg>
                  <span>เวลาที่แสดงเป็นการประมาณการณ์เท่านั้น และอาจเปลี่ยนแปลงได้ตามทรัพยากรของระบบ</span>
                </div>
              )}

              {progress === 100 && (
                <div className="mt-4 flex items-center justify-center text-green-600 bg-green-50 p-3 rounded-lg border border-green-200 animate-fadeIn">
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-2" viewBox="0 0 20 20" fill="currentColor">
                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                  </svg>
                  <span className="font-medium">การประมวลผลเสร็จสมบูรณ์แล้ว กำลังเตรียมการดาวน์โหลด...</span>
                </div>
              )}
            </div>
          </div>
        )}

        {error && (
          <div className="bg-red-50 text-red-800 p-4 rounded-lg border border-red-200 mb-6 flex items-start animate-fadeIn">
            <div className="flex-shrink-0 h-10 w-10 rounded-full bg-red-100 flex items-center justify-center mr-3">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-red-600" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7 4a1 1 0 11-2 0 1 1 0 012 0zm-1-9a1 1 0 00-1 1v4a1 1 0 102 0V6a1 1 0 00-1-1z" clipRule="evenodd" />
              </svg>
            </div>
            <div>
              <h3 className="text-sm font-medium text-red-800 mb-1">เกิดข้อผิดพลาด</h3>
              <p className="text-sm">{error}</p>
            </div>
          </div>
        )}

        {success && !isLoading && (
          <div className="bg-green-50 text-green-800 p-4 rounded-lg border border-green-200 mb-6 flex items-start animate-fadeIn">
            <div className="flex-shrink-0 h-10 w-10 rounded-full bg-green-100 flex items-center justify-center mr-3">
              <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 text-green-600" viewBox="0 0 20 20" fill="currentColor">
                <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
              </svg>
            </div>
            <div>
              <h3 className="text-sm font-medium text-green-800 mb-1">ดำเนินการเสร็จสิ้น</h3>
              <p className="text-sm whitespace-pre-line">{success}</p>
            </div>
          </div>
        )}

        <div className="text-center mb-8">
          <button
            onClick={handleGenerateCertificates}
            disabled={!excelFile || !templateFile || isLoading}
            className={`px-10 py-4 rounded-lg text-lg font-semibold shadow-lg transform transition-all duration-300 ${!excelFile || !templateFile || isLoading
                ? 'bg-gray-300 cursor-not-allowed'
                : 'bg-gradient-to-r from-blue-600 to-blue-500 text-white hover:from-blue-700 hover:to-blue-600 hover:shadow-xl hover:scale-105 active:scale-100'
              }`}
          >
            {isLoading ? (
              <div className="flex items-center justify-center">
                <svg className="animate-spin -ml-1 mr-3 h-5 w-5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                  <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                  <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                กำลังสร้าง...
              </div>
            ) : (
              <div className="flex items-center justify-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 mr-2" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
                </svg>
                สร้าง Certificate
              </div>
            )}
          </button>
          <p className="mt-2 text-sm text-gray-500">
            {!excelFile ? 'กรุณาอัพโหลดไฟล์ Excel ก่อน' :
              !templateFile ? 'กรุณาอัพโหลดไฟล์ Word Template ก่อน' :
                'พร้อมสร้าง Certificate แล้ว'}
          </p>
        </div>

        <div className="mb-8">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-gradient-to-r from-blue-50 to-indigo-50 p-6 rounded-xl border border-blue-200 shadow-sm">
              <h3 className="font-bold text-blue-800 text-lg mb-4 flex items-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 mr-2 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                วิธีใช้งาน
              </h3>
              <ol className="list-none space-y-4 text-gray-700 ml-2">
                <li className="relative pl-8">
                  <div className="absolute left-0 top-0 flex items-center justify-center h-6 w-6 rounded-full bg-blue-100 text-blue-800">1</div>
                  <div>
                    <span className="font-medium text-blue-700">อัพโหลดไฟล์ Excel</span>
                    <p className="text-sm mt-1 text-gray-600">อัพโหลดไฟล์ Excel ที่มีรายชื่อและข้อมูลของผู้รับ certificate</p>
                  </div>
                </li>
                <li className="relative pl-8">
                  <div className="absolute left-0 top-0 flex items-center justify-center h-6 w-6 rounded-full bg-blue-100 text-blue-800">2</div>
                  <div>
                    <span className="font-medium text-blue-700">อัพโหลดไฟล์ Word Template</span>
                    <p className="text-sm mt-1 text-gray-600">อัพโหลดไฟล์ Word ที่เป็น template โดยใส่ตัวแปรในรูปแบบ <code className="bg-blue-100 px-1.5 py-0.5 rounded text-blue-800 font-mono">{`{ชื่อตัวแปร}`}</code></p>
                  </div>
                </li>
                <li className="relative pl-8">
                  <div className="absolute left-0 top-0 flex items-center justify-center h-6 w-6 rounded-full bg-blue-100 text-blue-800">3</div>
                  <div>
                    <span className="font-medium text-blue-700">สร้าง Certificate</span>
                    <p className="text-sm mt-1 text-gray-600">กดปุ่ม สร้าง Certificate เพื่อสร้างและดาวน์โหลดไฟล์</p>
                  </div>
                </li>
              </ol>
            </div>

            <div className="bg-gradient-to-r from-purple-50 to-pink-50 p-6 rounded-xl border border-purple-200 shadow-sm">
              <h3 className="font-bold text-purple-800 text-lg mb-4 flex items-center">
                <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6 mr-2 text-purple-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z" />
                </svg>
                เครื่องมือเพิ่มเติม
              </h3>
              <div className="mb-4 p-4 bg-white rounded-lg border border-purple-200 transition-all hover:shadow-md duration-300">
                <h4 className="font-medium text-purple-800 mb-2">แปลงไฟล์ Word เป็น PDF</h4>
                <p className="text-sm text-gray-600 mb-3">
                  หลังจากสร้าง Certificate แล้ว คุณสามารถแปลงไฟล์ Word (.docx) เป็นไฟล์ PDF ทั้งหมดพร้อมกันได้ด้วยโค้ด VBA
                </p>
                <button
                  onClick={() => setShowVBAModal(true)}
                  className="w-full py-2 bg-gradient-to-r from-purple-600 to-purple-500 text-white rounded-md hover:from-purple-700 hover:to-purple-600 transition-colors font-medium flex items-center justify-center"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5 mr-1.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 20l4-16m4 4l4 4-4 4M6 16l-4-4 4-4" />
                  </svg>
                  ดูโค้ด VBA สำหรับแปลงเป็น PDF
                </button>
              </div>
              <div className="text-xs text-gray-500 mt-2 italic">
                * ต้องใช้ Microsoft Word ในการเปิดใช้งานโค้ด VBA
              </div>
            </div>
          </div>
        </div>

      
      </main>

      {/* Modal สำหรับโค้ด VBA */}
      <VBAModal />
    </div>
  );
}
