/*
 * msvc_crt_shim.c — helpers that WebView2LoaderStatic.lib needs.
 *
 * The static WebView2 loader (vendor/webview2-com-sys) ships as objects
 * compiled with MSVC; when linked on windows-gnu they reference a handful of
 * MSVC CRT symbols that neither the GNU toolchain nor mingw-w64 provide.
 * These are their minimal implementations; msvc_crt_shim.s adds the two
 * MSVC-mangled C++ names. build.rs compiles both only for windows-gnu.
 */
#include <windows.h>
#include <stdlib.h>

/* /GS stack cookie. MSVC startup fills this through _security_init_cookie,
 * but there is no MSVC CRT startup in a static link, so a fixed value is used
 * instead: every loader object reads and compares the same global, so the
 * integrity check still works. */
unsigned long long __security_cookie = 0x2B1A1A1BULL;

void __security_check_cookie(unsigned long long cookie) {
    if (cookie != __security_cookie)
        /* STATUS_STACK_BUFFER_OVERRUN: the same code MSVC's fastfail uses */
        RaiseException(0xC0000409UL, 0, 0, 0);
}

/* Thread-safe initialization of function-local statics, in the MSVC guard
 * protocol the precompiled loader objects expect: guard 0 = uninit,
 * -2 = some thread is initializing (the call-instruction sequence rechecks
 * for -2 after this function returns to decide who performs the init),
 * -1 = done. A Sleep loop is enough for the loader's one-shot initialization
 * (shortcut: recursive init from the owning thread would spin; the upstream
 * loader guards are never re-entered, revisit if a future loader version does).
 * _Init_thread_epoch is read by the compiled code but its value does not
 * matter for this implementation. */
volatile long _Init_thread_epoch = 0;

void _Init_thread_header(long* guard) {
    for (;;) {
        long v = *guard;
        if (v == -1)
            return; /* already initialized */
        if (v == 0 && InterlockedCompareExchange((volatile long*)guard, -2, 0) == 0)
            return; /* this thread now owns the init */
        while (*guard == -2)
            Sleep(5); /* another thread owns it: wait for the footer */
    }
}

void _Init_thread_footer(long* guard) { InterlockedExchange((volatile long*)guard, -1); }

/* std::nothrow new/delete. The loader only makes small allocations; returning
 * null on failure matches nothrow semantics exactly. */
void* isel_new_nothrow(unsigned long long size, void const* nothrow_tag) {
    (void)nothrow_tag;
    return malloc((size_t)size);
}

void isel_delete(void* ptr, unsigned long long size) {
    (void)size;
    free(ptr);
}
