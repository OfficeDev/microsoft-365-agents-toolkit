# Qt Code Review Guide

> Code review guidelines focusing on object model, signals/slots, event loop, and GUI performance. Examples based on Qt 5.15 / Qt 6.

## Table of Contents

- [Object Model & Memory Management](#object-model--memory-management)
- [Signals & Slots](#signals--slots)
- [Containers & Strings](#containers--strings)
- [Threads & Concurrency](#threads--concurrency)
- [GUI & Widgets](#gui--widgets)
- [Meta-Object System](#meta-object-system)
- [Review Checklist](#review-checklist)

---

## Object Model & Memory Management

### Use Parent-Child Ownership Mechanism
Qt's `QObject` hierarchy automatically manages memory. For `QObject`, prefer setting a parent object over manual `delete` or smart pointers.

```cpp
// ❌ Manual management prone to memory leaks
QWidget* w = new QWidget();
QLabel* l = new QLabel();
l->setParent(w);
// ... If w is deleted, l is automatically deleted. But if w leaks, l also leaks.

// ✅ Specify parent in constructor
QWidget* w = new QWidget(this); // Owned by 'this'
QLabel* l = new QLabel(w);      // Owned by 'w'
```

### Use Smart Pointers with QObject
If a `QObject` has no parent, use `QScopedPointer` or `std::unique_ptr` with a custom deleter (use `deleteLater` if cross-thread). Avoid `std::shared_ptr` for `QObject` unless necessary, as it confuses the parent-child ownership system.

```cpp
// ✅ Scoped pointer for local/member QObject without parent
QScopedPointer<MyObject> obj(new MyObject());

// ✅ Safe pointer to prevent dangling pointers
QPointer<MyObject> safePtr = obj.data();
if (safePtr) {
    safePtr->doSomething();
}
```

### Use `deleteLater()`
For asynchronous deletion, especially in slots or event handlers, use `deleteLater()` instead of `delete` to ensure pending events in the event loop are processed.

---

## Signals & Slots

### Prefer Function Pointer Syntax
Use compile-time checked syntax (Qt 5+).

```cpp
// ❌ String-based (runtime check only, slower)
connect(sender, SIGNAL(valueChanged(int)), receiver, SLOT(updateValue(int)));

// ✅ Compile-time check
connect(sender, &Sender::valueChanged, receiver, &Receiver::updateValue);
```

### Connection Types
Be explicit or aware of connection types when crossing threads.
- `Qt::AutoConnection` (Default): Direct if same thread, Queued if different thread.
- `Qt::QueuedConnection`: Always posts event (thread-safe across threads).
- `Qt::DirectConnection`: Immediate call (dangerous if accessing non-thread-safe data across threads).

### Avoid Loops
Check logic that might cause infinite signal loops (e.g., `valueChanged` -> `setValue` -> `valueChanged`). Block signals or check for equality before setting values.

```cpp
void MyClass::setValue(int v) {
    if (m_value == v) return; // ? Good: Break loop
    m_value = v;
    emit valueChanged(v);
}
```

---

## Containers & Strings

### QString Efficiency
- Use `QStringLiteral("...")` for compile-time string creation to avoid runtime allocation.
- Use `QLatin1String` for comparison with ASCII literals (in Qt 5).
- Prefer `arg()` for formatting (or `QStringBuilder`'s `%` operator).

```cpp
// ❌ Runtime conversion
if (str == "test") ...

// ✅ Prefer QLatin1String for comparison with ASCII literals (in Qt 5)
if (str == QLatin1String("test")) ... // Qt 5
if (str == u"test"_s) ...             // Qt 6
```

### Container Selection
- **Qt 6**: `QList` is now the default choice (unified with `QVector`).
- **Qt 5**: Prefer `QVector` over `QList` for contiguous memory and cache performance, unless stable references are needed.
- Be aware of Implicit Sharing (Copy-on-Write). Passing containers by value is cheap *until* modified. Use `const &` for read-only access.

```cpp
// ❌ Forces deep copy if function modifies 'list'
void process(QVector<int> list) {
    list[0] = 1; 
}

// ✅ Read-only reference
void process(const QVector<int>& list) { ... }
```

---

## Threads & Concurrency

### Subclassing QThread vs Worker Object
Prefer the "Worker Object" pattern over subclassing `QThread` implementation details.

```cpp
// ❌ Business logic inside QThread::run()
class MyThread : public QThread {
    void run() override { ... } 
};

// ✅ Worker object moved to thread
QThread* thread = new QThread;
Worker* worker = new Worker;
worker->moveToThread(thread);
connect(thread, &QThread::started, worker, &Worker::process);
thread->start();
```

### GUI Thread Safety
**NEVER** access UI widgets (`QWidget` and subclasses) from a background thread. Use signals/slots to communicate updates to the main thread.

---

## GUI & Widgets

### Logic Separation
Keep business logic out of UI classes (`MainWindow`, `Dialog`). UI classes should only handle display and user input forwarding.

### Layouts
Avoid fixed sizes (`setGeometry`, `resize`). Use layouts (`QVBoxLayout`, `QGridLayout`) to handle different DPIs and window resizing gracefully.

### Blocking Event Loop
Never execute long-running operations on the main thread (freezes GUI).
- **Bad**: `Sleep()`, `while(busy)`, synchronous network calls.
- **Good**: `QProcess`, `QThread`, `QtConcurrent`, or asynchronous APIs (`QNetworkAccessManager`).

---

## Meta-Object System

### Properties & Enums
Use `Q_PROPERTY` for values exposed to QML or needing introspection.
Use `Q_ENUM` to enable string conversion for enums.

```cpp
class MyObject : public QObject {
    Q_OBJECT
    Q_PROPERTY(int value READ value WRITE setValue NOTIFY valueChanged)
public:
    enum State { Idle, Running };
    Q_ENUM(State)
    // ...
};
```

### qobject_cast
Use `qobject_cast<T*>` for QObjects instead of `dynamic_cast`. It is faster and doesn't require RTTI.

---

## Review Checklist

- [ ] **Memory**: Is parent-child relationship correct? Are dangling pointers avoided (using `QPointer`)?
- [ ] **Signals**: Are connections checked? Do lambdas use safe captures (context object)?
- [ ] **Threads**: Is UI accessed only from main thread? Are long tasks offloaded?
- [ ] **Strings**: Are `QStringLiteral` or `tr()` used appropriately?
- [ ] **Style**: Naming conventions (camelCase for methods, PascalCase for classes).
- [ ] **Resources**: Are resources (images, styles) loaded from `.qrc`?