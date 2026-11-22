import { parseDownloadUrl, safeFilename } from "@/lib/download";
import { convertExcelToDocx } from "@/lib/excel-to-docx";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// GET handles streaming a remote report URL converted into Word (.docx)
export async function GET(request: Request) {
  const { searchParams } = new URL(request.url);

  const target = parseDownloadUrl(searchParams.get("url"));
  if (!target) {
    return Response.json({ success: false, message: "That download link is not allowed." }, { status: 400 });
  }

  let upstream: Response;
  try {
    upstream = await fetch(target.toString(), { cache: "no-store" });
  } catch {
    return Response.json({ success: false, message: "Could not fetch the report file." }, { status: 502 });
  }

  if (!upstream.ok || !upstream.body) {
    return Response.json(
      { success: false, message: "The report link has expired. Please generate it again." },
      { status: 502 }
    );
  }

  try {
    const arrayBuffer = await upstream.arrayBuffer();
    const rawFilename = searchParams.get("filename") || "report.xlsx";
    const docxFilename = safeFilename(rawFilename.replace(/\.xlsx$/i, ".docx"), "report.docx");

    const docxBytes = await convertExcelToDocx(arrayBuffer, {
      title: "Wint Wealth Report",
      orientation: "landscape",
    });

    const headers = new Headers({
      "Content-Type": DOCX_MIME,
      "Content-Disposition": `attachment; filename="${docxFilename}"`,
      "Content-Length": String(docxBytes.byteLength),
      "Cache-Control": "no-store",
    });

    return new Response(Buffer.from(docxBytes), { headers });
  } catch (err) {
    return Response.json(
      { success: false, message: "Failed to convert Excel to Word: " + (err instanceof Error ? err.message : String(err)) },
      { status: 500 }
    );
  }
}

// POST handles direct file upload conversion to Word (.docx)
export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof Blob)) {
      return Response.json({ success: false, message: "Please upload an Excel (.xlsx/.xls) file." }, { status: 400 });
    }

    const title = (formData.get("title") as string) || "Excel Report";
    const arrayBuffer = await file.arrayBuffer();

    const docxBytes = await convertExcelToDocx(arrayBuffer, {
      title,
      orientation: "landscape",
    });

    const originalName = file instanceof File ? file.name : "converted.xlsx";
    const docxFilename = safeFilename(originalName.replace(/\.(xlsx|xls|csv)$/i, ".docx"), "converted.docx");

    const headers = new Headers({
      "Content-Type": DOCX_MIME,
      "Content-Disposition": `attachment; filename="${docxFilename}"`,
      "Content-Length": String(docxBytes.byteLength),
      "Cache-Control": "no-store",
    });

    return new Response(Buffer.from(docxBytes), { headers });
  } catch (err) {
    return Response.json(
      { success: false, message: "Failed to convert file: " + (err instanceof Error ? err.message : String(err)) },
      { status: 500 }
    );
  }
}
