#include <execinfo.h>
#include <unistd.h>

#include <exception>

namespace {

[[noreturn]] void write_terminate_backtrace() noexcept {
    static constexpr char header[] = "XLSLIBERATOR_TERMINATE_BACKTRACE_BEGIN\n";
    static constexpr char footer[] = "XLSLIBERATOR_TERMINATE_BACKTRACE_END\n";
    static_cast<void>(write(STDERR_FILENO, header, sizeof(header) - 1));
    void* frames[128];
    const int count = backtrace(frames, 128);
    backtrace_symbols_fd(frames, count, STDERR_FILENO);
    static_cast<void>(write(STDERR_FILENO, footer, sizeof(footer) - 1));
    _exit(134);
}

__attribute__((constructor)) void install_terminate_backtrace() {
    std::set_terminate(write_terminate_backtrace);
}

}  // namespace
