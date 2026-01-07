import * as React from "react";
import { Spinner } from "@fluentui/react/lib/Spinner";
import * as XLSX from "xlsx";
import { useEffect, useState } from "react";
const MsgReader = require('@kenjiuno/msgreader').default;
interface DocumentViewerProps {
  url: string;
  isOpen: boolean;
  onDismiss: () => void;
  fileName: string;
  isLocalFile?: boolean;
}

const DocumentViewer: React.FC<DocumentViewerProps> = ({ url, isOpen, onDismiss, fileName }) => {
  const fileType = getFileType(fileName);
  const [officePreviewUrl, setOfficePreviewUrl] = useState<string | null>(null);
  const [excelHtml, setExcelHtml] = useState<string | null>(null);
  const [emlHtml, setEmlHtml] = useState<string | null>(null);
  const [msgHtml, setMsgHtml] = useState<string | null>(null);

  // Excel parsing effect
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

  //Office parsing effect
  useEffect(() => {
    if (fileType === "office" && url) {
      const serverRelativeUrl = url.replace(new RegExp(`^${window.location.origin}`), '');
      const previewUrl = `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(serverRelativeUrl)}`;
      setOfficePreviewUrl(previewUrl);
    } else {
      setOfficePreviewUrl(null);
    }
  }, [url, fileType]);

  // EML parsing effect
  useEffect(() => {
    if (fileType === 'eml' && url) {
      (async () => {
        try {
          const response = await fetch(url);
          const emlText = await response.text();

          // Regex to extract text/html part content from the raw EML string
          const match = emlText.match(/Content-Type: text\/html;[\s\S]*?([\s\S]*?)(?=--\w+)/i);
          if (match && match[1]) {
            // Simple cleanup of possible MIME encoding and boundary lines (basic)
            const cleanedHtml = match[1]
              .replace(/charset="?[\w\-]*"?/gi, "") // remove charset attribute inside HTML
              .replace(/^\s*charset=.*\n?/im, "") // also remove whole line if it's at the start
              .replace(/\r\n/g, '\n') // normalize line endings
              .replace(/^\s*<html>/i, '<div>') // optional: replace starting html tag for safer rendering
              .replace(/<\/html>\s*$/i, '</div>') // optional: close as div
              .replace(/^\s*P\.\{margin-top:0;margin-bottom:0;\}\s*$/im, "")
              .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, "") // Remove all <style>...</style> blocks
              .replace(/P\.\{margin-top:0;margin-bottom:0;\}/gi, ""); // Remove loose style text

            setEmlHtml(cleanedHtml);
          } else {
            setEmlHtml('<div>No HTML content found in this email.</div>');
          }
        } catch {
          setEmlHtml('<div>Failed to load EML file.</div>')
        }
      })();
    } else {
      setEmlHtml(null);
    }
  }, [url, fileType]);

  // MSG parsing effect
  useEffect(() => {
    const fetchAndParseMsg = async () => {
      if (fileType === "msg" && url) {
        try {
          const response = await fetch(url);
          const arrayBuffer = await response.arrayBuffer();
          const msgReader = new MsgReader(new Uint8Array(arrayBuffer));
          const msgData = msgReader.getFileData();

          let rawHtml = "";
          // Check if bodyHTML exists
          if (msgData.bodyHTML) {
            rawHtml = msgData.bodyHTML;
          }
          // If not, check for a Uint8Array HTML body
          else if (msgData.html && (msgData.html instanceof Uint8Array || Array.isArray(msgData.html))) {
            rawHtml = new TextDecoder('utf-8').decode(msgData.html);
          }

          // Clean up the HTML just as in your EML logic
          const cleanedHtml = rawHtml
            .replace(/charset="?[\w\-]*"?/gi, "") // Remove charset attribute
            .replace(/^\s*charset=.*\n?/im, "") // Remove charset line at start
            .replace(/\r\n/g, '\n') // Normalize line endings
            .replace(/^\s*<html>/i, '<div>') // Optional: for safer rendering
            .replace(/<\/html>\s*$/i, '</div>') // Optional: for safer rendering
            .replace(/^\s*P\.\{margin-top:0;margin-bottom:0;\}\s*$/im, "")
            .replace(/<style[\s\S]*?>[\s\S]*?<\/style>/gi, "") // Remove all <style></style> blocks
            .replace(/P\.\{margin-top:0;margin-bottom:0;\}/gi, "");


          setMsgHtml(cleanedHtml || msgData.body || "<div>No preview available</div>");
        } catch (error) {
          setMsgHtml("<div>Failed to load MSG file.</div>");
          console.log(error)
        }
      } else {
        setMsgHtml(null);
      }
    };

    fetchAndParseMsg();
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
            height="1000px"
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
      case "office":
        if (!officePreviewUrl) {
          return <Spinner label="Preparing Office preview..." />;
        }
        return (
          <iframe
            src={officePreviewUrl}
            width="100%"
            height="1000px"
            style={{ border: "none", minHeight: "600px" }}
            title={`${fileName} Preview`}
          />
        );
      case "eml":
        return emlHtml ? (
          <div
            dangerouslySetInnerHTML={{ __html: emlHtml }}
            style={{ overflow: "auto", maxHeight: "600px", textAlign: "left" }}
          />
        ) : (
          <Spinner label="Loading EML preview..." />
        );
      case "msg":
        return msgHtml ? (
          <div
            dangerouslySetInnerHTML={{ __html: msgHtml }}
            style={{ overflow: "auto", maxHeight: "600px", textAlign: "left" }}
          />
        ) : (
          <Spinner label="Loading MSG preview..." />
        );
      default:
        return <p>Unsupported file type</p>;
    }
  };

  return (
    url ? (
      <div style={{
        border: "2px dashed #ccc",
        borderRadius: 6,
        padding: 24,
        textAlign: "center",
        backgroundColor: "#fafafa",
        cursor: "pointer",
        color: "#999",
      }}>
        {renderContent()}
      </div>
    ) : (
      <div style={{
        border: "2px dashed #ccc",
        borderRadius: 6,
        padding: 24,
        textAlign: "center",
        backgroundColor: "#fafafa",
        cursor: "pointer",
        color: "#999",
      }}>
        <a>Uploaded attachments will be displayed here.</a>
      </div>
    )
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
  if (["eml"].includes(ext)) return "eml";
  if (["msg"].includes(ext)) return "msg";
  return "unknown";
}
