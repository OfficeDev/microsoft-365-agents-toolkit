# C++ Code Review Guide

> C++ code review guide focused on memory safety, lifetime, API design, and performance. Examples assume C++17/20.

## Table of Contents

- [Ownership and RAII](#ownership-and-raii)
- [Lifetime and References](#lifetime-and-references)
- [Copy and Move Semantics](#copy-and-move-semantics)
- [Const-Correctness and API Design](#const-correctness-and-api-design)
- [Error Handling and Exception Safety](#error-handling-and-exception-safety)
- [Concurrency](#concurrency)
- [Performance and Allocation](#performance-and-allocation)
- [Templates and Type Safety](#templates-and-type-safety)
- [Tooling and Build Checks](#tooling-and-build-checks)
- [Review Checklist](#review-checklist)

---

## Ownership and RAII

### Prefer RAII and smart pointers

Use RAII to express ownership. Default to `std::unique_ptr`, use `std::shared_ptr` only for shared lifetime.

```cpp
// ? Bad: manual new/delete with early returns
Foo* make_foo() {
    Foo* foo = new Foo();
    if (!foo->Init()) {
        delete foo;
        return nullptr;
    }
    return foo;
}

// ? Good: RAII with unique_ptr
std::unique_ptr<Foo> make_foo() {
    auto foo = std::make_unique<Foo>();
    if (!foo->Init()) {
        return {};
    }
    return foo;
}
```

### Wrap C resources

```cpp
// ? Good: wrap FILE* with unique_ptr
using FilePtr = std::unique_ptr<FILE, decltype(&fclose)>;

FilePtr open_file(const char* path) {
    return FilePtr(fopen(path, "rb"), &fclose);
}
```

---

## Lifetime and References

### Avoid dangling references and views

`std::string_view` and `std::span` do not own data. Make sure the owner outlives the view.

```cpp
// ? Bad: returning string_view to a temporary
std::string_view bad_view() {
    std::string s = make_name();
    return s; // dangling
}

// ? Good: return owning string
std::string good_name() {
    return make_name();
}

// ? Good: view tied to caller-owned data
std::string_view good_view(const std::string& s) {
    return s;
}
```

### Lambda captures

```cpp
// ? Bad: capture reference that escapes
std::function<void()> make_task() {
    int value = 42;
    return [&]() { use(value); }; // dangling
}

// ? Good: capture by value
std::function<void()> make_task() {
    int value = 42;
    return [value]() { use(value); };
}
```

---

## Copy and Move Semantics

### Rule of 0/3/5

Prefer the Rule of 0 by using RAII types. If you own a resource, define or delete copy and move operations.

```cpp
// ? Bad: raw ownership with default copy
struct Buffer {
    int* data;
    size_t size;
    explicit Buffer(size_t n) : data(new int[n]), size(n) {}
    ~Buffer() { delete[] data; }
    // copy ctor/assign are implicitly generated -> double delete
};

// ? Good: Rule of 0 with std::vector
struct Buffer {
    std::vector<int> data;
    explicit Buffer(size_t n) : data(n) {}
};
```

### Delete unwanted copies

```cpp
struct Socket {
    Socket() = default;
    ~Socket() { close(); }

    Socket(const Socket&) = delete;
    Socket& operator=(const Socket&) = delete;
    Socket(Socket&&) noexcept = default;
    Socket& operator=(Socket&&) noexcept = default;
};
```

---

## Const-Correctness and API Design

### Use const and explicit

```cpp
class User {
public:
    const std::string& name() const { return name_; }
    void set_name(std::string name) { name_ = std::move(name); }

private:
    std::string name_;
};

struct Millis {
    explicit Millis(int v) : value(v) {}
    int value;
};
```

### Avoid object slicing

```cpp
struct Shape { virtual ~Shape() = default; };
struct Circle : Shape { void draw() const; };

// ? Bad: slices Circle into Shape
void draw(Shape shape);

// ? Good: pass by reference
void draw(const Shape& shape);
```

### Use override and final

```cpp
struct Base {
    virtual void run() = 0;
};

struct Worker final : Base {
    void run() override {}
};
```

---

## Error Handling and Exception Safety

### Prefer RAII for cleanup

```cpp
// ? Good: RAII handles cleanup on exceptions
void process() {
    std::vector<int> data = load_data(); // safe cleanup
    do_work(data);
}
```

### Do not throw from destructors

```cpp
struct File {
    ~File() noexcept { close(); }
    void close();
};
```

### Use expected results for normal failures

```cpp
// ? Expected error: use optional or expected
std::optional<int> parse_int(const std::string& s) {
    try {
        return std::stoi(s);
    } catch (...) {
        return std::nullopt;
    }
}
```

---

## Concurrency

### Protect shared data

```cpp
// ? Bad: data race
int counter = 0;
void inc() { counter++; }

// ? Good: atomic
std::atomic<int> counter{0};
void inc() { counter.fetch_add(1, std::memory_order_relaxed); }
```

### Use RAII locks

```cpp
std::mutex mu;
std::vector<int> data;

void add(int v) {
    std::lock_guard<std::mutex> lock(mu);
    data.push_back(v);
}
```

---

## Performance and Allocation

### Avoid repeated allocations

```cpp
// ? Bad: repeated reallocation
std::vector<int> build(int n) {
    std::vector<int> out;
    for (int i = 0; i < n; ++i) {
        out.push_back(i);
    }
    return out;
}

// ? Good: reserve upfront
std::vector<int> build(int n) {
    std::vector<int> out;
    out.reserve(static_cast<size_t>(n));
    for (int i = 0; i < n; ++i) {
        out.push_back(i);
    }
    return out;
}
```

### String concatenation

```cpp
// ? Bad: repeated allocation
std::string join(const std::vector<std::string>& parts) {
    std::string out;
    for (const auto& p : parts) {
        out += p;
    }
    return out;
}

// ? Good: reserve total size
std::string join(const std::vector<std::string>& parts) {
    size_t total = 0;
    for (const auto& p : parts) {
        total += p.size();
    }
    std::string out;
    out.reserve(total);
    for (const auto& p : parts) {
        out += p;
    }
    return out;
}
```

---

## Templates and Type Safety

### Prefer constrained templates (C++20)

```cpp
// ? Bad: overly generic
template <typename T>
T add(T a, T b) {
    return a + b;
}

// ? Good: constrained
template <typename T>
requires std::is_integral_v<T>
T add(T a, T b) {
    return a + b;
}
```

### Use static_assert for invariants

```cpp
template <typename T>
struct Packet {
    static_assert(std::is_trivially_copyable_v<T>,
        "Packet payload must be trivially copyable");
    T payload;
};
```

---

## Tooling and Build Checks

```bash
# Warnings
clang++ -Wall -Wextra -Werror -Wconversion -Wshadow -std=c++20 ...

# Sanitizers (debug builds)
clang++ -fsanitize=address,undefined -fno-omit-frame-pointer -g ...
clang++ -fsanitize=thread -fno-omit-frame-pointer -g ...

# Static analysis
clang-tidy src/*.cpp -- -std=c++20

# Formatting
clang-format -i src/*.cpp include/*.h
```

---

## Review Checklist

### Safety and Lifetime
- [ ] Ownership is explicit (RAII, unique_ptr by default)
- [ ] No dangling references or views
- [ ] Rule of 0/3/5 followed for resource-owning types
- [ ] No raw new/delete in business logic
- [ ] Destructors are noexcept and do not throw

### API and Design
- [ ] const-correctness is applied consistently
- [ ] Constructors are explicit where needed
- [ ] Override/final used for virtual functions
- [ ] No object slicing (pass by ref or pointer)

### Concurrency
- [ ] Shared data is protected (mutex or atomics)
- [ ] Locking order is consistent
- [ ] No blocking while holding locks

### Performance
- [ ] Unnecessary allocations avoided (reserve, move)
- [ ] Copies avoided in hot paths
- [ ] Algorithmic complexity is reasonable

### Tooling and Tests
- [ ] Builds clean with warnings enabled
- [ ] Sanitizers run on critical code paths
- [ ] Static analysis (clang-tidy) results are addressed
