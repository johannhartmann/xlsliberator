#include <execinfo.h>
#include <unistd.h>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <exception>
#include <typeinfo>

namespace {

void write_text(const char* text) noexcept {
    static_cast<void>(write(STDERR_FILENO, text, std::strlen(text)));
}

void write_backtrace(const char* header, const char* footer) noexcept {
    static thread_local bool writing = false;
    if (writing) {
        return;
    }
    writing = true;
    write_text(header);
    void* frames[64];
    const int count = backtrace(frames, 64);
    backtrace_symbols_fd(frames, count, STDERR_FILENO);
    write_text(footer);
    writing = false;
}

void write_current_exception() noexcept {
    try {
        const std::exception_ptr current = std::current_exception();
        if (!current) {
            write_text("exception_type=<none>\n");
            return;
        }
        std::rethrow_exception(current);
    } catch (const std::exception& exc) {
        char message[1024];
        const int length = std::snprintf(
            message,
            sizeof(message),
            "exception_type=%s exception_what=%s\n",
            typeid(exc).name(),
            exc.what()
        );
        if (length > 0) {
            message[sizeof(message) - 1] = '\0';
            write_text(message);
        }
    } catch (...) {
        write_text("exception_type=<non-std-exception>\n");
    }
}

[[noreturn]] void write_terminate_backtrace() noexcept {
    write_text("XLSLIBERATOR_TERMINATE_EXCEPTION_BEGIN\n");
    write_current_exception();
    write_text("XLSLIBERATOR_TERMINATE_EXCEPTION_END\n");
    write_backtrace(
        "XLSLIBERATOR_TERMINATE_BACKTRACE_BEGIN\n",
        "XLSLIBERATOR_TERMINATE_BACKTRACE_END\n"
    );
    _exit(134);
}

__attribute__((constructor)) void install_terminate_backtrace() {
    std::set_terminate(write_terminate_backtrace);
}

}  // namespace

extern "C" [[noreturn]] void abort() noexcept {
    write_backtrace(
        "XLSLIBERATOR_ABORT_BACKTRACE_BEGIN\n",
        "XLSLIBERATOR_ABORT_BACKTRACE_END\n"
    );
    _exit(134);
}
