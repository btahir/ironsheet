import { OoxmlPackage, Workbook, type CompressionAdapter } from "@ironsheet/core";

export const browserCompressionAdapter: CompressionAdapter = {
  async inflateRaw(data) {
    return pipeBytesThroughCompressionStream(data, "deflate-raw", "decompress");
  },
  async deflateRaw(data) {
    return pipeBytesThroughCompressionStream(data, "deflate-raw", "compress");
  }
};

export async function openWorkbookFromBlob(blob: Blob): Promise<Workbook> {
  return openWorkbookFromArrayBuffer(await blob.arrayBuffer());
}

export async function openWorkbookFromArrayBuffer(data: ArrayBuffer): Promise<Workbook> {
  return openWorkbookFromBytes(new Uint8Array(data));
}

export async function openWorkbookFromBytes(data: Uint8Array): Promise<Workbook> {
  return Workbook.fromPackage(OoxmlPackage.open(data, browserCompressionAdapter));
}

export async function writeWorkbookToBlob(
  workbook: Workbook,
  options: { type?: string } = {}
): Promise<Blob> {
  const data = await workbook.write();
  return new Blob([bytesToArrayBuffer(data)], {
    type: options.type ?? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}

async function pipeBytesThroughCompressionStream(
  data: Uint8Array,
  format: CompressionFormat,
  mode: "compress" | "decompress"
): Promise<Uint8Array> {
  const stream =
    mode === "compress" ? new CompressionStream(format) : new DecompressionStream(format);
  const writer = stream.writable.getWriter();
  const outputPromise = readStreamBytes(stream.readable);
  await writer.write(new Uint8Array(bytesToArrayBuffer(data)));
  await writer.close();

  return outputPromise;
}

async function readStreamBytes(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let totalBytes = 0;

  while (true) {
    const result = await reader.read();
    if (result.done) {
      break;
    }

    chunks.push(result.value);
    totalBytes += result.value.byteLength;
  }

  const output = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }

  return output;
}

function bytesToArrayBuffer(data: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(data.byteLength);
  copy.set(data);
  return copy.buffer;
}
