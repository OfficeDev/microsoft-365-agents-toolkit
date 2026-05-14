# Scenarios

End-to-end user scenarios. Each scenario reads from a [persona](../personas.md), runs through the lifecycle, and surfaces the exact toolkit features in play.

These scenarios feed [UX flows](../ux/flows/README.md) (one-to-one mapping where it makes sense).

## Pages

- [build-declarative-agent.md](build-declarative-agent.md)
- [build-custom-engine-agent.md](build-custom-engine-agent.md)
- [build-bot.md](build-bot.md)
- [build-message-extension.md](build-message-extension.md)
- [build-graph-connector.md](build-graph-connector.md)

## Scenario template

```markdown
# <Scenario name>

**Persona:** P1 / P2 / ...
**Outcome:** <one sentence>
**Surface:** VS Code · Visual Studio · CLI

## Steps
1. Scaffold — template ID, language, key questions
2. Run / debug — local target (Playground / Teams sideload / etc.)
3. Provision — drivers invoked, prerequisites
4. Deploy — code package format, target service
5. Publish — destination (org catalog / store)

## Files produced
- ...

## Drivers in play
- ...

## Where this is tested
- ...
```
