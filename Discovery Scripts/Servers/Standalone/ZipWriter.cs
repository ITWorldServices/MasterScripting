// Small ZIP container writer for .NET 2.0+. No external assemblies or binaries.
using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Collections;
using System.Collections.Generic;

public static class StandaloneDiscoveryZip {
    private sealed class Entry {
        public byte[] Name;
        public uint Crc, Size, CompressedSize, Offset;
    }
    private static uint Crc32(byte[] data) {
        uint[] table = new uint[256];
        for (uint i = 0; i < 256; i++) {
            uint c = i;
            for (int bit = 0; bit < 8; bit++) c = (c & 1) != 0 ? 0xedb88320U ^ (c >> 1) : c >> 1;
            table[i] = c;
        }
        uint crc = 0xffffffffU;
        foreach (byte value in data) crc = table[(crc ^ value) & 255] ^ (crc >> 8);
        return crc ^ 0xffffffffU;
    }
    public static void Write(string path, IDictionary parts) {
        if (parts.Count > 65535) throw new InvalidOperationException("Too many XLSX parts.");
        List<Entry> entries = new List<Entry>();
        using (BinaryWriter writer = new BinaryWriter(new FileStream(path, FileMode.CreateNew))) {
            foreach (DictionaryEntry item in parts) {
                byte[] data = new UTF8Encoding(false).GetBytes((string)item.Value);
                byte[] packed;
                using (MemoryStream buffer = new MemoryStream()) {
                    using (DeflateStream deflater = new DeflateStream(buffer, CompressionMode.Compress)) {
                        deflater.Write(data, 0, data.Length);
                    }
                    packed = buffer.ToArray();
                }
                if (writer.BaseStream.Position + packed.Length + 1024 > UInt32.MaxValue)
                    throw new InvalidOperationException("Workbook exceeds the ZIP32 limit; the CLIXML snapshot retains the inventory.");
                Entry entry = new Entry();
                entry.Name = Encoding.UTF8.GetBytes((string)item.Key);
                entry.Crc = Crc32(data); entry.Size = (uint)data.Length;
                entry.CompressedSize = (uint)packed.Length; entry.Offset = (uint)writer.BaseStream.Position;
                writer.Write(0x04034b50U); writer.Write((ushort)20); writer.Write((ushort)0x800);
                writer.Write((ushort)8); writer.Write((ushort)0); writer.Write((ushort)0x2821);
                writer.Write(entry.Crc); writer.Write(entry.CompressedSize); writer.Write(entry.Size);
                writer.Write((ushort)entry.Name.Length); writer.Write((ushort)0); writer.Write(entry.Name); writer.Write(packed);
                entries.Add(entry);
            }
            uint start = (uint)writer.BaseStream.Position;
            foreach (Entry entry in entries) {
                writer.Write(0x02014b50U); writer.Write((ushort)20); writer.Write((ushort)20);
                writer.Write((ushort)0x800); writer.Write((ushort)8); writer.Write((ushort)0); writer.Write((ushort)0x2821);
                writer.Write(entry.Crc); writer.Write(entry.CompressedSize); writer.Write(entry.Size);
                writer.Write((ushort)entry.Name.Length); writer.Write((ushort)0); writer.Write((ushort)0);
                writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(0U); writer.Write(entry.Offset); writer.Write(entry.Name);
            }
            uint size = (uint)writer.BaseStream.Position - start;
            writer.Write(0x06054b50U); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write((ushort)entries.Count); writer.Write((ushort)entries.Count);
            writer.Write(size); writer.Write(start); writer.Write((ushort)0);
        }
    }
}
