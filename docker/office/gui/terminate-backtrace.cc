#include <execinfo.h>
#include <unistd.h>

#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <exception>
#include <new>

namespace {

void write_backtrace(const char* header, const char* footer) noexcept {
    static thread_local bool writing = false;
    if (writing) {
        return;
    }
    writing = true;
    static_cast<void>(write(STDERR_FILENO, header, std::strlen(header)));
    void* frames[64];
    const int count = backtrace(frames, 64);
    backtrace_symbols_fd(frames, count, STDERR_FILENO);
    static_cast<void>(write(STDERR_FILENO, footer, std::strlen(footer)));
    writing = false;
}

[[noreturn]] void write_terminate_backtrace() noexcept {
    write_backtrace(
        "XLSLIBERATOR_TERMINATE_BACKTRACE_BEGIN\n",
        "XLSLIBERATOR_TERMINATE_BACKTRACE_END\n"
    );
    _exit(134);
}

[[noreturn]] void throw_allocation_failure(std::size_t size) {
    char header[128];
    std::snprintf(
        header,
        sizeof(header),
        "XLSLIBERATOR_ALLOCATION_FAILURE_BEGIN size=%zu\n",
        size
    );
    write_backtrace(header, "XLSLIBERATOR_ALLOCATION_FAILURE_END\n");
    throw std::bad_alloc();
}

void* allocate(std::size_t size) {
    if (void* memory = std::malloc(size)) {
        return memory;
    }
    throw_allocation_failure(size);
}

__attribute__((constructor)) void install_terminate_backtrace() {
    std::set_terminate(write_terminate_backtrace);
}

}  // namespace

void* operator new(std::size_t size) {
    return allocate(size);
}

void* operator new[](std::size_t size) {
    return allocate(size);
}

void operator delete(void* memory) noexcept {
    std::free(memory);
}

void operator delete[](void* memory) noexcept {
    std::free(memory);
}

void operator delete(void* memory, std::size_t) noexcept {
    std::free(memory);
}

void operator delete[](void* memory, std::size_t) noexcept {
    std::free(memory);
}
