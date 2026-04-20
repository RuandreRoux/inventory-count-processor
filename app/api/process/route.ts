import { NextRequest, NextResponse } from "next/server";
import * as XLSX from "xlsx";
import { processWorkbook } from "@/lib/excel-processor";

export const runtime = "nodejs";
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File | null;

    if (!file) {
      return NextResponse.json({ error: "No file provided" }, { status: 400 });
    }

    const ext = file.name.split(".").pop()?.toLowerCase();
    if (!["xls", "xlsx", "xlsm"].includes(ext ?? "")) {
      return NextResponse.json(
        { error: "Please upload an Excel file (.xls or .xlsx)" },
        { status: 400 }
      );
    }

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    const { workbook, stats } = processWorkbook(buffer);

    const outputBuffer = XLSX.write(workbook, {
      type: "buffer",
      bookType: "xlsx",
    });

    const date = new Date().toISOString().slice(0, 10);
    const filename = `inventory-count-cleaned-${date}.xlsx`;

    return new NextResponse(outputBuffer, {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "X-Stats": JSON.stringify(stats),
      },
    });
  } catch (err) {
    console.error("Processing error:", err);
    return NextResponse.json(
      { error: "Failed to process file. Make sure it is a valid Sage Evolution export." },
      { status: 500 }
    );
  }
}
