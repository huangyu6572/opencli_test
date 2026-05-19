# opencli-plugin-opencli-plugin-devcloud

Huawei DevCloud sprint adapter

## Install

```bash
# From local development directory
opencli plugin install file://D:\code\opencli_test\plugin-scaffold

# From GitHub (after publishing)
opencli plugin install github:<user>/opencli-plugin-opencli-plugin-devcloud
```

## Commands

| Command | Type | Description |
|---------|------|-------------|
| `opencli-plugin-devcloud/hello` | Pipeline | Sample pipeline command |
| `opencli-plugin-devcloud/greet` | TypeScript | Sample TS command with func() |

## Development

```bash
# Install locally for development (symlinked, changes reflect immediately)
opencli plugin install file://D:\code\opencli_test\plugin-scaffold

# Verify commands are registered
opencli list | grep opencli-plugin-devcloud

# Run a command
opencli opencli-plugin-devcloud hello
opencli opencli-plugin-devcloud greet --name World
```
