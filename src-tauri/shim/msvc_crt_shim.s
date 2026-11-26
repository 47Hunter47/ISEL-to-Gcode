/*
 * The MSVC-mangled names baked into the precompiled WebView2 loader objects.
 * x64 has one calling convention, so a tail jmp preserves the arguments.
 * mingw-w64 C symbols carry no leading underscore on x64.
 */
    .section .text
    .globl "??2@YAPEAX_KAEBUnothrow_t@std@@@Z"
"??2@YAPEAX_KAEBUnothrow_t@std@@@Z":
    jmp isel_new_nothrow

    .globl "??3@YAXPEAX_K@Z"
"??3@YAXPEAX_K@Z":
    jmp isel_delete

    .section .rdata
    .globl "?nothrow@std@@3Unothrow_t@1@B"
"?nothrow@std@@3Unothrow_t@1@B":
    .byte 0
