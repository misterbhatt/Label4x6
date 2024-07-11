function reloadPage() {
    location.reload();
}

document.getElementById('fileInput').addEventListener('change', handleFile, false);

let excelData = [];

function handleFile(e) {
    console.log("File input changed");
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function(event) {
        console.log("File loaded successfully");
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        excelData = XLSX.utils.sheet_to_json(worksheet);
        document.getElementById('errorMessage').textContent = '';
        console.log("Excel data parsed successfully:", excelData);
    };

    reader.onerror = function(event) {
        const errorMessage = 'Error reading file';
        console.error(errorMessage);
        document.getElementById('errorMessage').textContent = errorMessage;
    };

    reader.readAsArrayBuffer(file);
}

function generatePDF() {
    console.log("Generate PDF Labels button clicked");
    if (!excelData.length) {
        const errorMessage = 'Please upload a valid Excel file';
        console.error(errorMessage);
        document.getElementById('errorMessage').textContent = errorMessage;
        return;
    }

    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({
            orientation: 'portrait',
            unit: 'in',
            format: [4, 6]
        });

        excelData.forEach((row, index) => {
            doc.setFontSize(9);
            doc.setFont("helvetica", "bold");
            doc.text('Receiver Information:-', 0.2, 0.3);

            doc.setFont("helvetica", "bold");
            let yPos = 0.5;
            doc.text(`Receiver Name: ${row['Receiver Name'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            doc.text(`Phone: ${row['Receiver Phone'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            const receiverAddressLine1 = doc.splitTextToSize(`Address: ${row['Receiver Address Line 1'] || ''}`, 3.6);
            doc.text(receiverAddressLine1, 0.2, yPos);
            yPos += 0.2 * receiverAddressLine1.length;
            const receiverAddressLine2 = doc.splitTextToSize(row['Receiver Address Line 2'] || '', 3.6);
            doc.text(receiverAddressLine2, 0.2, yPos);
            yPos += 0.2 * receiverAddressLine2.length;
            doc.text(`City: ${row['City'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            doc.text(`State: ${row['State'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            doc.text(`Pincode: ${row['Receiver Pincode'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            
            // Spacer line after Shipment Details
            doc.setLineWidth(0.01);  // Set line width
            doc.line(0.2, yPos, 3.8, yPos);  // Draw line
            yPos += 0.1;  // Adjust the vertical position after the line

            // product details
            doc.setFontSize(9);
            doc.setFont("helvetica", "bold");
            doc.text('Shipment Details:-', 0.2, yPos);
            doc.setFont("helvetica", "bold");
            yPos += 0.3;
            doc.text(`AWB No: ${row['Airwaybill Number'] || ''}`, 0.2, yPos);
            JsBarcode("#barcode", row['Airwaybill Number'] || '', {format: "CODE128"});
            doc.addImage(document.getElementById('barcode').toDataURL(), 'PNG', 2, yPos - 0.2, 1.6, 0.4);
            yPos += 0.3;
            
            doc.setFontSize(8);
            doc.setFont("helvetica", "bold");
            doc.text('SKU:-', 0.2, yPos);
            doc.setFont("helvetica", "normal");
            yPos += 0.2;
            const SKUNumber = doc.splitTextToSize(`${row['SKU Number'] || ''}`, 3.6);
            doc.text(SKUNumber, 0.2, yPos);
            yPos += 0.2 * SKUNumber.length;

            doc.setFontSize(8);
            doc.setFont("helvetica", "normal");
            //yPos += 0.2;
            doc.text(`Weight: ${row['Weight'] || ''}`, 0.2, yPos);
            doc.text(`Quantity: ${row['Quantity'] || ''}`, 2.0, yPos);  // Adjust the x position for Quantity
            yPos += 0.2;
            doc.text(`Amount: ${row['Amount'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            
            doc.setFontSize(8);
            doc.setFont("helvetica", "bold");
            doc.text('Description:-', 0.2, yPos);
            doc.setFont("helvetica", "normal");
            yPos += 0.2;
            const productDescription = doc.splitTextToSize(`${row['Product Description'] || ''}`, 3.6);
            doc.text(productDescription, 0.2, yPos);
            yPos += 0.2 * productDescription.length;
            doc.text(`Order No: ${row['Reference Number'] || ''}`, 0.2, yPos);
            JsBarcode("#barcodeRef", row['Reference Number'] || '', {format: "CODE128"});
            doc.addImage(document.getElementById('barcodeRef').toDataURL(), 'PNG', 2, yPos - 0.2, 1.6, 0.4);
            yPos += 0.3;

            // Spacer line after Shipment Details
            doc.setLineWidth(0.01);  // Set line width
            doc.line(0.2, yPos, 3.8, yPos);  // Draw line
            yPos += 0.1;  // Adjust the vertical position after the line

            // sender information details

            doc.setFontSize(8);
            doc.setFont("helvetica", "bold");
            doc.text('Return Address:-', 0.2, yPos); //Sender Information
            doc.setFont("helvetica", "normal");
            yPos += 0.2;
            doc.text(`Name: ${row['Sender Name'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            doc.text(`Phone: ${row['Sender Phone'] || ''}`, 0.2, yPos);
            yPos += 0.2;
            const senderAddressLine1 = doc.splitTextToSize(`Address: ${row['Sender Address Line 1'] || ''}`, 3.6)
            doc.text(senderAddressLine1, 0.2, yPos);
            yPos += 0.2 * senderAddressLine1.length;
            const senderAddressLine2 = doc.splitTextToSize(row['Sender Address Line 2'] || '', 3.6);
            doc.text(senderAddressLine2, 0.2, yPos);
            yPos += 0.2 * senderAddressLine2.length;
            doc.text(`Pincode: ${row['Sender Pincode'] || ''}`, 0.2, yPos);
            
            if (index < excelData.length - 1) {
                doc.addPage();
            }
        });

        doc.save('Labels.pdf');
        console.log("PDF Labels generated successfully");
        document.getElementById('errorMessage').textContent = 'PDF Labels generated successfully';
    } catch (error) {
        console.error("Error generating PDF:", error);
        document.getElementById('errorMessage').textContent = 'Error generating PDF';
    }
}

function downloadSampleFile() {
    console.log("Download Sample File button clicked");
    const sampleData = [
        {
            "Reference Number": "REF123",
            "Airwaybill Number": "AWB123456",
            "Receiver Name": "John Doe",
            "Receiver Phone": "1234567890",
            "Receiver Address Line 1": "123 Main St",
            "Receiver Address Line 2": "",
            "City": "Delhi",
            "State":"Delhi",
            "Receiver Pincode": "123456",
            "SKU Number": "SKU123",
            "Weight": "0.5",
            "Quantity": "1",
            "Amount": "100",
            "Product Description": "Sample Product",
            "Sender Name": "Jane Smith",
            "Sender Phone": "0987654321",
            "Sender Address Line 1": "456 Another St",
            "Sender Address Line 2": "",
            "Sender Pincode": "654321"
        }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sample Data");

    XLSX.writeFile(workbook, 'SampleData.xlsx');
    console.log("Sample file downloaded successfully");
}
