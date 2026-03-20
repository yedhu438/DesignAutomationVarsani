def _compress_channel_zip(channel_bytes: bytes) -> bytes:
    """
    Raw deflate for PSD compression mode 2 (ZIP without prediction).
    PSD spec requires raw deflate — NO zlib wrapper (no 0x789C header,
    no Adler-32 checksum). zlib.compress() adds both; wbits=-15 strips them.
    """
    obj = zlib.compressobj(level=6, method=zlib.DEFLATED, wbits=-15)
    return obj.compress(channel_bytes) + obj.flush()

