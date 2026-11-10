import { parseDownloadUrl, safeFilename } from "@/lib/download";
import { convertExcelToPdf } from "@/lib/excel-to-pdf";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

// GET handles streaming a remote report URL converted into PDF
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
    const pdfFilename = safeFilename(rawFilename.replace(/\.xlsx$/i, ".pdf"), "report.pdf");

    const { pdfBytes } = convertExcelToPdf(arrayBuffer, {
      title: "Wint Wealth Report",
      subtitle: `Generated on ${new Date().toLocaleDateString()}`,
      orientation: "auto",
    });

    const headers = new Headers({
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${pdfFilename}"`,
      "Content-Length": String(pdfBytes.byteLength),
      "Cache-Control": "no-store",
    });

    return new Response(Buffer.from(pdfBytes), { headers });
  } catch (err) {
    return Response.json(
      { success: false, message: "Failed to convert Excel to PDF: " + (err instanceof Error ? err.message : String(err)) },
      { status: 500 }
    );
  }
}

// POST handles direct file upload conversion
export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");

    if (!file || !(file instanceof Blob)) {
      return Response.json({ success: false, message: "Please upload an Excel (.xlsx/.xls) file." }, { status: 400 });
    }

    const orientation = (formData.get("orientation") as "portrait" | "landscape" | "auto") || "auto";
    const title = (formData.get("title") as string) || "Excel Report";
    const arrayBuffer = await file.arrayBuffer();

    const { pdfBytes } = convertExcelToPdf(arrayBuffer, {
      title,
      orientation,
    });

    const originalName = file instanceof File ? file.name : "converted.xlsx";
    const pdfFilename = safeFilename(originalName.replace(/\.(xlsx|xls|csv)$/i, ".pdf"), "converted.pdf");

    const headers = new Headers({
      "Content-Type": "application/pdf",
      "Content-Disposition": `attachment; filename="${pdfFilename}"`,
      "Content-Length": String(pdfBytes.byteLength),
      "Cache-Control": "no-store",
    });

    return new Response(Buffer.from(pdfBytes), { headers });
  } catch (err) {
    return Response.json(
      { success: false, message: "Failed to convert file: " + (err instanceof Error ? err.message : String(err)) },
      { status: 500 }
    );
  }
}
