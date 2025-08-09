import React, { useState } from "react";
import ExcelJS from "exceljs";

export default function HtmlToXlsxConverter() {
    const [xlsxBlob, setXlsxBlob] = useState(null);
    const [tableData, setTableData] = useState([]);

    const handleFileUpload = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (event) => {
            const htmlString = event.target.result;

            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlString, "text/html");
            const table = doc.querySelector("table");

            if (!table) {
                alert("No table found in the uploaded HTML file.");
                return;
            }

            const previewData = Array.from(table.rows).map((row) =>
                Array.from(row.cells).map((cell) => cell.textContent.trim())
            );
            setTableData(previewData);

            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet("Sheet1");

            previewData.forEach((rowData, rowIndex) => {
                const row = worksheet.addRow(rowData);
                row.eachCell((cell) => {
                    cell.alignment = { horizontal: "right" };
                    if (rowIndex === 0) {
                        cell.font = { bold: true };
                    }
                });
            });

            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], {
                type:
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            setXlsxBlob(blob);
        };
        reader.readAsText(file);
    };

    const handleDownload = () => {
        if (!xlsxBlob) return;
        const url = URL.createObjectURL(xlsxBlob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "converted.xlsx";
        a.click();
        URL.revokeObjectURL(url);
    };

    // visually-hidden style for screen-reader-only text
    const srOnly = {
        position: "absolute",
        width: "1px",
        height: "1px",
        padding: 0,
        margin: "-1px",
        overflow: "hidden",
        clip: "rect(0,0,0,0)",
        whiteSpace: "nowrap",
        border: 0,
    };

    return (
        <div
            style={{
                fontFamily:
                    "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
                backgroundColor: "#f2f2f7",
                minHeight: "95vh",
                display: "flex",
                justifyContent: "center",
                padding: "20px",
            }}
        >
            <div
                role="region"
                aria-labelledby="converter-title"
                style={{
                    backgroundColor: "#fff",
                    borderRadius: "20px",
                    boxShadow: "0 4px 20px rgba(0,0,0,0.1)",
                    padding: "20px",
                    width: "100%",
                }}
            >
                {/* Header Message */}
                <div
                    style={{
                        textAlign: "center",
                        marginBottom: "20px",
                        borderBottom: "1px solid #e5e5ea",
                        paddingBottom: "10px",
                    }}
                >
                    <h2
                        id="converter-title"
                        style={{
                            fontSize: "22px",
                            fontWeight: "600",
                            margin: "0",
                            color: "#000",
                        }}
                    >
                        HTML to XLSX Converter
                    </h2>
                    <p
                        style={{
                            fontSize: "15px",
                            marginTop: "6px",
                            color: "#3c3c43",
                        }}
                    >
                        Upload an HTML file containing a table to preview and download as
                        Excel.
                    </p>
                </div>

                {/* Top bar with file upload & download */}
                <div
                    style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        marginBottom: "20px",
                    }}
                >
                    {/* Upload label acts as button */}
                    <label
                        htmlFor="file-upload"
                        style={{
                            display: "inline-flex",
                            alignItems: "center",
                            gap: "8px",
                            padding: "10px 16px",
                            backgroundColor: "#007aff",
                            color: "#fff",
                            borderRadius: "12px",
                            fontSize: "16px",
                            fontWeight: "500",
                            cursor: "pointer",
                            transition: "background-color 0.3s",
                        }}
                        onMouseEnter={(e) =>
                            (e.currentTarget.style.backgroundColor = "#005bbb")
                        }
                        onMouseLeave={(e) =>
                            (e.currentTarget.style.backgroundColor = "#007aff")
                        }
                        aria-label="Upload HTML file"
                        title="Upload HTML file"
                    >
                        {/* Upload SVG icon (arrow up into tray) */}
                        <svg xmlns="http://www.w3.org/2000/svg" fill="#ffffff" width="24px" height="24px" viewBox="0 0 24 24"><path d="M9 16h6v-6h4l-7-7-7 7h4zm-4 2h14v2H5z" /></svg>

                        <span>Choose HTML File</span>
                        <span style={srOnly}>Choose an HTML or HTM file to upload</span>
                    </label>

                    <input
                        id="file-upload"
                        type="file"
                        accept=".html,.htm"
                        onChange={handleFileUpload}
                        style={{ display: "none" }}
                    />

                    {/* Download button with SVG */}
                    {xlsxBlob && (
                        <button
                            onClick={handleDownload}
                            style={{
                                display: "inline-flex",
                                alignItems: "center",
                                gap: "8px",
                                padding: "10px 16px",
                                backgroundColor: "#34c759",
                                color: "#fff",
                                border: "none",
                                borderRadius: "12px",
                                fontSize: "16px",
                                fontWeight: "500",
                                cursor: "pointer",
                                transition: "background-color 0.3s",
                            }}
                            onMouseEnter={(e) =>
                                (e.currentTarget.style.backgroundColor = "#28a745")
                            }
                            onMouseLeave={(e) =>
                                (e.currentTarget.style.backgroundColor = "#34c759")
                            }
                            aria-label="Download converted XLSX file"
                            title="Download XLSX"
                        >
                            {/* Download SVG icon (arrow down into tray) */}
                            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>

                            <span>Download XLSX</span>
                        </button>
                    )}
                </div>

                {/* Preview Table */}
                {tableData.length > 0 && (
                    <div
                        style={{
                            overflowX: "auto",
                            overflowY: "auto",
                            scrollbarWidth: "thin",
                            maxHeight: "70vh",
                            borderRadius: "12px",
                            border: "1px solid #d1d1d6",
                        }}
                    >
                        <table
                            style={{
                                borderCollapse: "collapse",
                                width: "100%",
                                minWidth: "600px",
                                backgroundColor: "#fff",
                            }}
                        >
                            <thead
                                style={{
                                    backgroundColor: "#f2f2f7",
                                }}
                            >
                                <tr>
                                    {tableData[0].map((header, i) => (
                                        <th
                                            key={i}
                                            style={{
                                                padding: "12px",
                                                fontWeight: "600",
                                                borderBottom: "1px solid #d1d1d6",
                                                position: "sticky",
                                                top: 0,
                                                background: "#f2f2f7",
                                            }}
                                        >
                                            {header}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {tableData.slice(1).map((row, rIdx) => (
                                    <tr key={rIdx}>
                                        {row.map((cell, cIdx) => (
                                            <td
                                                key={cIdx}
                                                style={{
                                                    padding: "12px",
                                                    borderBottom: "1px solid #e5e5ea",
                                                }}
                                            >
                                                {cell}
                                            </td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>
        </div>
    );
}
