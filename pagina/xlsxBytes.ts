// NFXLSX1 carries an already generated XLSX. Never parse or regenerate its cells.
export const MAX_XLSX_BYTES = 30 * 1024 * 1024;
export const MAX_COMPRESSED_XLSX_BYTES = MAX_XLSX_BYTES + 65536;

export function validarXlsxBytes(bytes: Uint8Array): Uint8Array {
    if (bytes.length > MAX_XLSX_BYTES) {
        throw new Error("O Excel ultrapassa 30 MB. Nenhum arquivo parcial foi baixado.");
    }
    // XLSX is a ZIP package beginning with a local-file header.
    if (bytes.length < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b || bytes[2] !== 0x03 || bytes[3] !== 0x04) {
        throw new Error("O texto recebido não contém a assinatura ZIP de um arquivo XLSX.");
    }
    return bytes;
}

export async function inflarXlsxComLimite(compressed: Uint8Array): Promise<Uint8Array> {
    if (compressed.length > MAX_COMPRESSED_XLSX_BYTES) {
        throw new Error("O arquivo recebido ultrapassa 30 MB.");
    }
    const stream = new Blob([compressed as BlobPart]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    const reader = stream.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    try {
        while (true) {
            const part = await reader.read();
            if (part.done) { break; }
            total += part.value.length;
            if (total > MAX_XLSX_BYTES) {
                await reader.cancel();
                throw new Error("O Excel ultrapassa 30 MB. Nenhum arquivo parcial foi baixado.");
            }
            chunks.push(part.value);
        }
    } finally {
        reader.releaseLock();
    }
    const bytes = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    return validarXlsxBytes(bytes);
}
