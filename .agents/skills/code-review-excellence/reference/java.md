# Java Code Review Guide

Java 审查重点：Java 17/21 新特性、Spring Boot 3 最佳实践、并发编程（虚拟线程）、JPA 性能优化以及代码可维护性。

## 目录

- [现代 Java 特性 (17/21+)](#现代-java-特性-1721)
- [Stream API & Optional](#stream-api--optional)
- [Spring Boot 最佳实践](#spring-boot-最佳实践)
- [JPA 与 数据库性能](#jpa-与-数据库性能)
- [并发与虚拟线程](#并发与虚拟线程)
- [Lombok 使用规范](#lombok-使用规范)
- [异常处理](#异常处理)
- [测试规范](#测试规范)
- [Review Checklist](#review-checklist)

---

## 现代 Java 特性 (17/21+)

### Record (记录类)

```java
// ❌ 传统的 POJO/DTO：样板代码多
public class UserDto {
    private final String name;
    private final int age;

    public UserDto(String name, int age) {
        this.name = name;
        this.age = age;
    }
    // getters, equals, hashCode, toString...
}

// ✅ 使用 Record：简洁、不可变、语义清晰
public record UserDto(String name, int age) {
    // 紧凑构造函数进行验证
    public UserDto {
        if (age < 0) throw new IllegalArgumentException("Age cannot be negative");
    }
}
```

### Switch 表达式与模式匹配

```java
// ❌ 传统的 Switch：容易漏掉 break，不仅冗长且易错
String type = "";
switch (obj) {
    case Integer i: // Java 16+
        type = String.format("int %d", i);
        break;
    case String s:
        type = String.format("string %s", s);
        break;
    default:
        type = "unknown";
}

// ✅ Switch 表达式：无穿透风险，强制返回值
String type = switch (obj) {
    case Integer i -> "int %d".formatted(i);
    case String s  -> "string %s".formatted(s);
    case null      -> "null value"; // Java 21 处理 null
    default        -> "unknown";
};
```

### 文本块 (Text Blocks)

```java
// ❌ 拼接 SQL/JSON 字符串
String json = "{\n" +
              "  \"name\": \"Alice\",\n" +
              "  \"age\": 20\n" +
              "}";

// ✅ 使用文本块：所见即所得
String json = """
    {
      "name": "Alice",
      "age": 20
    }
    """;
```

---

## Stream API & Optional

### 避免滥用 Stream

```java
// ❌ 简单的循环不需要 Stream（性能开销 + 可读性差）
items.stream().forEach(item -> {
    process(item);
});

// ✅ 简单场景直接用 for-each
for (var item : items) {
    process(item);
}

// ❌ 极其复杂的 Stream 链
List<Dto> result = list.stream()
    .filter(...)
    .map(...)
    .peek(...)
    .sorted(...)
    .collect(...); // 难以调试

// ✅ 拆分为有意义的步骤
var filtered = list.stream().filter(...).toList();
// ...
```

### Optional 正确用法

```java
// ❌ 将 Optional 用作参数或字段（序列化问题，增加调用复杂度）
public void process(Optional<String> name) { ... }
public class User {
    private Optional<String> email; // 不推荐
}

// ✅ Optional 仅用于返回值
public Optional<User> findUser(String id) { ... }

// ❌ 既然用了 Optional 还在用 isPresent() + get()
Optional<User> userOpt = findUser(id);
if (userOpt.isPresent()) {
    return userOpt.get().getName();
} else {
    return "Unknown";
}

// ✅ 使用函数式 API
return findUser(id)
    .map(User::getName)
    .orElse("Unknown");
```

---

## Spring Boot 最佳实践

### 依赖注入 (DI)

```java
// ❌ 字段注入 (@Autowired)
// 缺点：难以测试（需要反射注入），掩盖了依赖过多的问题，且不可变性差
@Service
public class UserService {
    @Autowired
    private UserRepository userRepo;
}

// ✅ 构造器注入 (Constructor Injection)
// 优点：依赖明确，易于单元测试 (Mock)，字段可为 final
@Service
public class UserService {
    private final UserRepository userRepo;

    public UserService(UserRepository userRepo) {
        this.userRepo = userRepo;
    }
}
// 💡 提示：结合 Lombok @RequiredArgsConstructor 可简化代码，但要小心循环依赖
```

### 配置管理

```java
// ❌ 硬编码配置值
@Service
public class PaymentService {
    private String apiKey = "sk_live_12345";
}

// ❌ 直接使用 @Value 散落在代码中
@Value("${app.payment.api-key}")
private String apiKey;

// ✅ 使用 @ConfigurationProperties 类型安全配置
@ConfigurationProperties(prefix = "app.payment")
public record PaymentProperties(String apiKey, int timeout, String url) {}
```

---

## JPA 与 数据库性能

### N+1 查询问题

```java
// ❌ FetchType.EAGER 或 循环中触发懒加载
// Entity 定义
@Entity
public class User {
    @OneToMany(fetch = FetchType.EAGER) // 危险！
    private List<Order> orders;
}

// 业务代码
List<User> users = userRepo.findAll(); // 1 条 SQL
for (User user : users) {
    // 如果是 Lazy，这里会触发 N 条 SQL
    System.out.println(user.getOrders().size());
}

// ✅ 使用 @EntityGraph 或 JOIN FETCH
@Query("SELECT u FROM User u JOIN FETCH u.orders")
List<User> findAllWithOrders();
```

### 事务管理

```java
// ❌ 在 Controller 层开启事务（数据库连接占用时间过长）
// ❌ 在 private 方法上加 @Transactional（AOP 不生效）
@Transactional
private void saveInternal() { ... }

// ✅ 在 Service 层公共方法加 @Transactional
// ✅ 读操作显式标记 readOnly = true (性能优化)
@Service
public class UserService {
    @Transactional(readOnly = true)
    public User getUser(Long id) { ... }

    @Transactional
    public void createUser(UserDto dto) { ... }
}
```

### Entity 设计

```java
// ❌ 在 Entity 中使用 Lombok @Data
// @Data 生成的 equals/hashCode 包含所有字段，可能触发懒加载导致性能问题或异常
@Entity
@Data
public class User { ... }

// ✅ 仅使用 @Getter, @Setter
// ✅ 自定义 equals/hashCode (通常基于 ID)
@Entity
@Getter
@Setter
public class User {
    @Id
    private Long id;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof User)) return false;
        return id != null && id.equals(((User) o).id);
    }

    @Override
    public int hashCode() {
        return getClass().hashCode();
    }
}
```

---

## 并发与虚拟线程

### 虚拟线程 (Java 21+)

```java
// ❌ 传统线程池处理大量 I/O 阻塞任务（资源耗尽）
ExecutorService executor = Executors.newFixedThreadPool(100);

// ✅ 使用虚拟线程处理 I/O 密集型任务（高吞吐量）
// Spring Boot 3.2+ 开启：spring.threads.virtual.enabled=true
ExecutorService executor = Executors.newVirtualThreadPerTaskExecutor();

// 在虚拟线程中，阻塞操作（如 DB 查询、HTTP 请求）几乎不消耗 OS 线程资源
```

### 线程安全

```java
// ❌ SimpleDateFormat 是线程不安全的
private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

// ✅ 使用 DateTimeFormatter (Java 8+)
private static final DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd");

// ❌ HashMap 在多线程环境可能死循环或数据丢失
// ✅ 使用 ConcurrentHashMap
Map<String, String> cache = new ConcurrentHashMap<>();
```

---

## Lombok 使用规范

```java
// ❌ 滥用 @Builder 导致无法强制校验必填字段
@Builder
public class Order {
    private String id; // 必填
    private String note; // 选填
}
// 调用者可能漏掉 id: Order.builder().note("hi").build();

// ✅ 关键业务对象建议手动编写 Builder 或构造函数以确保不变量
// 或者在 build() 方法中添加校验逻辑 (Lombok @Builder.Default 等)
```

---

## 异常处理

### 全局异常处理

```java
// ❌ 到处 try-catch 吞掉异常或只打印日志
try {
    userService.create(user);
} catch (Exception e) {
    e.printStackTrace(); // 不应该在生产环境使用
    // return null; // 吞掉异常，上层不知道发生了什么
}

// ✅ 自定义异常 + @ControllerAdvice (Spring Boot 3 ProblemDetail)
public class UserNotFoundException extends RuntimeException { ... }

@RestControllerAdvice
public class GlobalExceptionHandler {
    @ExceptionHandler(UserNotFoundException.class)
    public ProblemDetail handleNotFound(UserNotFoundException e) {
        return ProblemDetail.forStatusAndDetail(HttpStatus.NOT_FOUND, e.getMessage());
    }
}
```

---

## 测试规范

### 单元测试 vs 集成测试

```java
// ❌ 单元测试依赖真实数据库或外部服务
@SpringBootTest // 启动整个 Context，慢
public class UserServiceTest { ... }

// ✅ 单元测试使用 Mockito
@ExtendWith(MockitoExtension.class)
class UserServiceTest {
    @Mock UserRepository repo;
    @InjectMocks UserService service;

    @Test
    void shouldCreateUser() { ... }
}

// ✅ 集成测试使用 Testcontainers
@Testcontainers
@SpringBootTest
class UserRepositoryTest {
    @Container
    static PostgreSQLContainer<?> postgres = new PostgreSQLContainer<>("postgres:15");
    // ...
}
```

---

## Review Checklist

### 基础与规范
- [ ] 遵循 Java 17/21 新特性（Switch 表达式, Records, 文本块）
- [ ] 避免使用已过时的类（Date, Calendar, SimpleDateFormat）
- [ ] 集合操作是否优先使用了 Stream API 或 Collections 方法？
- [ ] Optional 仅用于返回值，未用于字段或参数

### Spring Boot
- [ ] 使用构造器注入而非 @Autowired 字段注入
- [ ] 配置属性使用了 @ConfigurationProperties
- [ ] Controller 职责单一，业务逻辑下沉到 Service
- [ ] 全局异常处理使用了 @ControllerAdvice / ProblemDetail

### 数据库 & 事务
- [ ] 读操作事务标记了 `@Transactional(readOnly = true)`
- [ ] 检查是否存在 N+1 查询（EAGER fetch 或循环调用）
- [ ] Entity 类未使用 @Data，正确实现了 equals/hashCode
- [ ] 数据库索引是否覆盖了查询条件

### 并发与性能
- [ ] I/O 密集型任务是否考虑了虚拟线程？
- [ ] 线程安全类是否使用正确（ConcurrentHashMap vs HashMap）
- [ ] 锁的粒度是否合理？避免在锁内进行 I/O 操作

### 可维护性
- [ ] 关键业务逻辑有充分的单元测试
- [ ] 日志记录恰当（使用 Slf4j，避免 System.out）
- [ ] 魔法值提取为常量或枚举
