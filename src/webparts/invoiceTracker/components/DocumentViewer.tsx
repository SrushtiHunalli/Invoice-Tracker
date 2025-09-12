/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
// import { Panel, PanelType } from "@fluentui/react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import * as XLSX from "xlsx";
import { useEffect, useState } from "react";

interface DocumentViewerProps {
  url: string;
  isOpen: boolean;
  onDismiss: () => void;
  fileName: string;
}

const DocumentViewer: React.FC<DocumentViewerProps> = ({ url, isOpen, onDismiss, fileName }) => {
  const fileType = getFileType(fileName);
  // Assuming the url is the relative path, encode it for use in the SharePoint preview URL
//   const encodedRelativePath = encodeURIComponent(url);
//   const previewUrl = `https://sachagroup.sharepoint.com/sites/Stagingsales/_layouts/15/Doc.aspx?sourcedoc=${encodedRelativePath}&file=${fileName}&action=embedview`;

  const [excelHtml, setExcelHtml] = useState<string | null>(null);

  useEffect(() => {
    const fetchAndParseExcel = async () => {
      if ((fileType === "excel" || fileType === "office") && url) {
        try {
          const response = await fetch(url);
          const arrayBuffer = await response.arrayBuffer();
          const workbook = XLSX.read(arrayBuffer, { type: "array" });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const html = XLSX.utils.sheet_to_html(worksheet);
          setExcelHtml(html);
        } catch (error) {
          setExcelHtml("<div>Failed to load Excel file.</div>");
        }
      } else {
        setExcelHtml(null);
      }
    };

    fetchAndParseExcel();
  }, [url, fileType]);

  const renderContent = () => {
    switch (fileType) {
      case "image":
        return <img src={url} alt="attachment" style={{ maxWidth: "100%", height: "auto" }} />;
      case "pdf":
        return (
          <iframe
            src={url}
            width="100%"
            height="600px"
            style={{ border: "none" }}
            title="PDF Viewer"
          />
        );
      case "excel":
        return excelHtml ? (
          <div
            dangerouslySetInnerHTML={{ __html: excelHtml }}
            style={{ overflow: "auto", maxHeight: "600px" }}
          />
        ) : (
          <Spinner label="Loading Excel..." />
        );
    //   case "office":
        
      default:
        return <p>Unsupported file type</p>;
    }
  };

  return (
    // <Panel
    //   isOpen={isOpen}
    //   onDismiss={onDismiss}
    //   headerText="Document Viewer"
    //   type={PanelType.large}
    //   closeButtonAriaLabel="Close"
    // >
      url ? <div style={{
        border: "2px dashed #ccc",
        borderRadius: 6,
        padding: 24,
        textAlign: "center",
        backgroundColor: "#fafafa",
        cursor: "pointer",
        color: "#999",
      }}>{renderContent()}</div> : <div style={{
        border: "2px dashed #ccc",
        borderRadius: 6,
        padding: 24,
        textAlign: "center",
        backgroundColor: "#fafafa",
        cursor: "pointer",
        color: "#999",
      }}><a>Uploaded attachments will be displayed here.</a></div>
    //   {/*
    // </Panel>
  );
};

export default DocumentViewer;

function getFileType(fileName: string): string {
  const ext = fileName.split('.').pop()?.toLowerCase();
  if (!ext) return "unknown";
  if (["jpg", "jpeg", "png", "gif", "bmp", "webp"].includes(ext)) return "image";
  if (["pdf"].includes(ext)) return "pdf";
  if (["xls", "xlsx", "csv"].includes(ext)) return "excel";
  if (["doc", "docx", "ppt", "pptx"].includes(ext)) return "office";
  return "unknown";
}
