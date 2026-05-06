export function readUInt16LE(data: Uint8Array, offset: number): number {
  return byteAt(data, offset) | (byteAt(data, offset + 1) << 8);
}

export function readUInt32LE(data: Uint8Array, offset: number): number {
  return (
    (byteAt(data, offset) |
      (byteAt(data, offset + 1) << 8) |
      (byteAt(data, offset + 2) << 16) |
      (byteAt(data, offset + 3) << 24)) >>>
    0
  );
}

export function readUInt64LE(data: Uint8Array, offset: number): bigint {
  return BigInt(readUInt32LE(data, offset)) | (BigInt(readUInt32LE(data, offset + 4)) << 32n);
}

export function writeUInt16LE(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
}

export function writeUInt32LE(target: Uint8Array, offset: number, value: number): void {
  target[offset] = value & 0xff;
  target[offset + 1] = (value >>> 8) & 0xff;
  target[offset + 2] = (value >>> 16) & 0xff;
  target[offset + 3] = (value >>> 24) & 0xff;
}

export function concatBytes(chunks: Uint8Array[]): Uint8Array {
  const size = chunks.reduce((total, chunk) => total + chunk.byteLength, 0);
  const output = new Uint8Array(size);
  let offset = 0;

  for (const chunk of chunks) {
    output.set(chunk, offset);
    offset += chunk.byteLength;
  }

  return output;
}

function byteAt(data: Uint8Array, offset: number): number {
  const byte = data[offset];
  if (byte === undefined) {
    throw new RangeError(`Unexpected end of data at byte ${offset}`);
  }

  return byte;
}
