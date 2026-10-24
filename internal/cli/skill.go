package cli

import (
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/spf13/cobra"
)

// The skill follows the Agent Skills format (agentskills.io): a folder named
// after it, holding SKILL.md (frontmatter, then instructions) and
// references/ with one file per topic of the guide, which SKILL.md tells
// the agent when to read. Clients load the frontmatter at startup, SKILL.md
// when a task needs it, and a reference when the instructions point there.

const skillName = "sheets"

// skillFront is SKILL.md's frontmatter. It holds only the specification's
// fields, so the skill loads in every client (claude.ai uploads refuse any
// other key), and plain one-line values (no ": ", " #" or "---" in them),
// which every YAML parser reads alike.
const skillFront = `---
name: sheets
description: Use this skill whenever a task involves an Excel workbook (.xlsx, .xlsm) or a CSV or TSV file - finding out what a spreadsheet holds, answering questions from its data, adding or changing rows, values or formulas, inserting or deleting rows, columns or sheets, sorting, copying, moving or find-and-replace, notes and hyperlinks, charts, making a sheet look professional or styling it as asked (number formats, highlights, dropdowns, frozen headers), converting between CSV and Excel, or checking a workbook for errors before handing it over. Use it even when the user only names the file or calls it the report or the budget, and use it instead of writing openpyxl or pandas code. It runs the sheets command-line tool, whose output is sized for the context window. Not for Google Sheets, Apple Numbers or legacy .xls files.
license: MIT
compatibility: Needs the sheets command-line tool on the PATH (go install github.com/eemax/sheets/cmd/sheets@latest). Works on local files, without network access.
metadata:
  author: eemax
  repository: https://github.com/eemax/sheets
---

`

// skillBody is the instructions with the guide's topics linked by link
// (each [[topic]] of skill-body.md).
func skillBody(link func(topic string) string) string {
	body := skillBodySource
	for _, t := range docTopics {
		body = strings.ReplaceAll(body, "[["+t.name+"]]", link(t.name))
	}
	return body
}

// skillFiles are the skill's files, by path within its folder.
func skillFiles() map[string]string {
	files := map[string]string{
		"SKILL.md": skillFront + skillBody(func(t string) string {
			return "[references/" + t + ".md](references/" + t + ".md)"
		}),
	}
	sections := docSections()
	for _, t := range docTopics {
		// A section stands alone as a file: its heading at the top.
		files["references/"+t.name+".md"] = "#" + strings.TrimPrefix(sections[t.heading], "##")
	}
	return files
}

// installSkill writes the skill into a skills directory, replacing what an
// earlier install wrote there, and returns its folder.
func installSkill(dir string) (string, error) {
	root := filepath.Join(dir, skillName)
	files := skillFiles()
	// References of topics the guide no longer has go.
	if entries, err := os.ReadDir(filepath.Join(root, "references")); err == nil {
		for _, e := range entries {
			if _, ours := files["references/"+e.Name()]; !ours && !e.IsDir() && strings.HasSuffix(e.Name(), ".md") {
				if err := os.Remove(filepath.Join(root, "references", e.Name())); err != nil {
					return root, err
				}
			}
		}
	}
	names := make([]string, 0, len(files))
	for name := range files {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		path := filepath.Join(root, filepath.FromSlash(name))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			return root, err
		}
		if err := os.WriteFile(path, []byte(files[name]), 0o644); err != nil {
			return root, err
		}
	}
	return root, nil
}

func (a *app) skillCmd() *cobra.Command {
	var install, project, agents bool
	var dir string
	cmd := &cobra.Command{
		Use:   "skill",
		Short: "Print or install the agent skill that teaches agents to use sheets",
		Long: `Print the sheets skill's SKILL.md, or install the skill so agents pick sheets
up by themselves when a task involves a spreadsheet. The skill follows the
Agent Skills format (agentskills.io): a sheets folder with SKILL.md and
references/, one file per topic of the guide.

  sheets skill --install                  # ~/.claude/skills/sheets (Claude Code)
  sheets skill --install --agents         # ~/.agents/skills/sheets (VS Code, Codex and other agents)
  sheets skill --install --project        # ./.claude/skills/sheets, this project only
  sheets skill --install --dir DIR        # DIR/sheets

Install it again after updating sheets, so the skill matches the tool.`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if !install {
				if project || agents || dir != "" {
					return fmt.Errorf("--project, --agents and --dir go with --install")
				}
				a.printf("%s", skillFiles()["SKILL.md"])
				return nil
			}
			if dir == "" {
				client := ".claude"
				if agents {
					client = ".agents"
				}
				base := "."
				if !project {
					home, err := os.UserHomeDir()
					if err != nil {
						return err
					}
					base = home
				}
				dir = filepath.Join(base, client, "skills")
			} else if project || agents {
				return fmt.Errorf("--dir names the skills directory itself; leave out --project and --agents")
			}
			root := filepath.Join(dir, skillName)
			refs := len(skillFiles()) - 1
			if a.dryRun {
				a.printDryRun()
				a.printf("would install %s (SKILL.md and %d references)\n", root, refs)
				return nil
			}
			root, err := installSkill(dir)
			if err != nil {
				return err
			}
			a.printf("installed %s (SKILL.md and %d references)\n", root, refs)
			return nil
		},
	}
	f := cmd.Flags()
	f.BoolVar(&install, "install", false, "write the skill folder instead of printing SKILL.md")
	f.BoolVar(&agents, "agents", false, "install into .agents/skills, which VS Code, Codex and other agents read, instead of .claude/skills")
	f.BoolVar(&project, "project", false, "install into the current directory instead of your home directory")
	f.StringVar(&dir, "dir", "", "install into this skills directory")
	return cmd
}
