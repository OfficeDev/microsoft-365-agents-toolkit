# C Code Review Guide

> C code review guide focused on memory safety, undefined behavior, and portability. Examples assume C11.

## Table of Contents

- [Pointer and Buffer Safety](#pointer-and-buffer-safety)
- [Ownership and Resource Management](#ownership-and-resource-management)
- [Undefined Behavior Pitfalls](#undefined-behavior-pitfalls)
- [Integer Types and Overflow](#integer-types-and-overflow)
- [Error Handling](#error-handling)
- [Concurrency](#concurrency)
- [Macros and Preprocessor](#macros-and-preprocessor)
- [API Design and Const](#api-design-and-const)
- [Tooling and Build Checks](#tooling-and-build-checks)
- [Review Checklist](#review-checklist)

---

## Pointer and Buffer Safety

### Always carry size with buffers

```c
// ? Bad: ignores destination size
bool copy_name(char *dst, size_t dst_size, const char *src) {
    strcpy(dst, src);
    return true;
}

// ? Good: validate size and terminate
bool copy_name(char *dst, size_t dst_size, const char *src) {
    size_t len = strlen(src);
    if (len + 1 > dst_size) {
        return false;
    }
    memcpy(dst, src, len + 1);
    return true;
}
```

### Avoid dangerous APIs

Prefer `snprintf`, `fgets`, and explicit bounds over `gets`, `strcpy`, or `sprintf`.

```c
// ? Bad: unbounded write
sprintf(buf, "%s", input);

// ? Good: bounded write
snprintf(buf, buf_size, "%s", input);
```

### Use the right copy primitive

```c
// ? Bad: memcpy with overlapping regions
memcpy(dst, src, len);

// ? Good: memmove handles overlap
memmove(dst, src, len);
```

---

## Ownership and Resource Management

### One allocation, one free

Track ownership and clean up on every error path.

```c
// ? Good: cleanup label avoids leaks
int load_file(const char *path) {
    int rc = -1;
    FILE *f = NULL;
    char *buf = NULL;

    f = fopen(path, "rb");
    if (!f) {
        goto cleanup;
    }
    buf = malloc(4096);
    if (!buf) {
        goto cleanup;
    }

    if (fread(buf, 1, 4096, f) == 0) {
        goto cleanup;
    }

    rc = 0;

cleanup:
    free(buf);
    if (f) {
        fclose(f);
    }
    return rc;
}
```

---

## Undefined Behavior Pitfalls

### Common UB patterns

```c
// ? Bad: use after free
char *p = malloc(10);
free(p);
p[0] = 'a';

// ? Bad: uninitialized read
int x;
if (x > 0) { /* UB */ }

// ? Bad: signed overflow
int sum = a + b;
```

### Avoid pointer arithmetic past the object

```c
// ? Bad: pointer past the end then dereference
int arr[4];
int *p = arr + 4;
int v = *p; // UB
```

---

## Integer Types and Overflow

### Avoid signed/unsigned surprises

```c
// ? Bad: negative converted to large size_t
int len = -1;
size_t n = len;

// ? Good: validate before converting
if (len < 0) {
    return -1;
}
size_t n = (size_t)len;
```

### Check for overflow in size calculations

```c
// ? Bad: potential overflow in multiplication
size_t bytes = count * sizeof(Item);

// ? Good: check before multiplying
if (count > SIZE_MAX / sizeof(Item)) {
    return NULL;
}
size_t bytes = count * sizeof(Item);
```

---

## Error Handling

### Always check return values

```c
// ? Bad: ignore errors
fread(buf, 1, size, f);

// ? Good: handle errors
size_t read = fread(buf, 1, size, f);
if (read != size && ferror(f)) {
    return -1;
}
```

### Consistent error contracts

- Use a clear convention: 0 for success, negative for failure.
- Document ownership rules on success and failure.
- If using `errno`, set it only for actual failures.

---

## Concurrency

### volatile is not synchronization

```c
// ? Bad: data race
volatile int stop = 0;
void worker(void) {
    while (!stop) { /* ... */ }
}

// ? Good: C11 atomics
_Atomic int stop = 0;
void worker(void) {
    while (!atomic_load(&stop)) { /* ... */ }
}
```

### Use mutexes for shared state

Protect shared data with `pthread_mutex_t` or equivalent. Avoid holding locks while doing I/O.

---

## Macros and Preprocessor

### Parenthesize arguments

```c
// ? Bad: macro with side effects
#define MIN(a, b) ((a) < (b) ? (a) : (b))
int x = MIN(i++, j++);

// ? Good: static inline function
static inline int min_int(int a, int b) {
    return a < b ? a : b;
}
```

---

## API Design and Const

### Const-correctness and sizes

```c
// ? Good: explicit size and const input
int hash_bytes(const uint8_t *data, size_t len, uint8_t *out);
```

### Document nullability

Clearly document whether pointers may be NULL. Prefer returning error codes instead of NULL when possible.

---

## Tooling and Build Checks

```bash
# Warnings
clang -Wall -Wextra -Werror -Wconversion -Wshadow -std=c11 ...

# Sanitizers (debug builds)
clang -fsanitize=address,undefined -fno-omit-frame-pointer -g ...
clang -fsanitize=thread -fno-omit-frame-pointer -g ...

# Static analysis
clang-tidy src/*.c -- -std=c11
cppcheck --enable=warning,performance,portability src/

# Formatting
clang-format -i src/*.c include/*.h
```

---

## Review Checklist

### Memory and UB
- [ ] All buffers have explicit size parameters
- [ ] No out-of-bounds access or pointer arithmetic past objects
- [ ] No use after free or uninitialized reads
- [ ] Signed overflow and shift rules are respected

### API and Design
- [ ] Ownership rules are documented and consistent
- [ ] const-correctness is applied for inputs
- [ ] Error contracts are clear and consistent

### Concurrency
- [ ] No data races on shared state
- [ ] volatile is not used for synchronization
- [ ] Locks are held for minimal time

### Tooling and Tests
- [ ] Builds clean with warnings enabled
- [ ] Sanitizers run on critical code paths
- [ ] Static analysis results are addressed
