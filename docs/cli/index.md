# CLI reference

Static command reference for the `ezxl` executable.

## 💻 Usage

```bash
ezxl [OPTIONS] COMMAND [ARGS]...
```

## ⚙️ Global options

| Option      | Short | Description                               |
| :---------- | :---- | :---------------------------------------- |
| `--help`    | `-h`  | Show the top-level help message and exit. |
| `--version` | `-v`  | Show the CLI version and exit.            |

## 📋 Commands

| Command   | Description                                                 |
| :-------- | :---------------------------------------------------------- |
| `version` | Display package version information.                        |
| `info`    | Display installed package metadata and dependency versions. |
| `docs`    | Open the online documentation in the default browser.       |

### `version`

| Option   | Short | Description                           |
| :------- | :---- | :------------------------------------ |
| `--full` | `-f`  | Display extended version information. |

### `info`

No command-specific options.

### `docs`

No command-specific options.

## 🧪 Examples

```bash
ezxl --version
ezxl version --full
ezxl info
ezxl docs
```
