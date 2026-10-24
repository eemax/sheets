package cli

import (
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
)

const skillFront = `---
name: sheets
description: Read, write and format Excel spreadsheets (.xlsx, .xlsm) and CSV files with the sheets CLI. Use whenever a task involves a spreadsheet or CSV file - seeing what it holds, reading or searching cells, writing data or formulas, sorting, copying, moving or replacing cells, inserting or deleting rows, columns or sheets, notes and links, making a sheet look professional, or styling it as asked (fills, number formats, highlights, dropdowns, frozen panes), or converting between CSV and Excel. Prefer it over writing openpyxl or pandas code; its output is sized for your context.
---

`

// skill returns the Claude Code skill: the guide's index with frontmatter.
// The index is short, and points at the guide's topics (sheets agent-docs
// TOPIC), so a task loads only the parts it needs.
func skill() string { return skillFront + agentIndex }

func (a *app) skillCmd() *cobra.Command {
	var install, project bool
	var dir string
	cmd := &cobra.Command{
		Use:   "skill",
		Short: "Print or install the Claude Code skill that teaches agents to use sheets",
		Long: `Print the sheets skill for Claude Code (SKILL.md), or install it so agents
pick sheets up by themselves when a task involves a spreadsheet:

  sheets skill --install            # ~/.claude/skills/sheets, all projects
  sheets skill --install --project  # ./.claude/skills/sheets, this project`,
		Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			if !install {
				a.printf("%s", skill())
				return nil
			}
			if dir == "" {
				if project {
					dir = filepath.Join(".claude", "skills")
				} else {
					home, err := os.UserHomeDir()
					if err != nil {
						return err
					}
					dir = filepath.Join(home, ".claude", "skills")
				}
			}
			target := filepath.Join(dir, "sheets", "SKILL.md")
			if a.dryRun {
				a.printDryRun()
				a.printf("would write %s\n", target)
				return nil
			}
			if err := os.MkdirAll(filepath.Dir(target), 0o755); err != nil {
				return err
			}
			if err := os.WriteFile(target, []byte(skill()), 0o644); err != nil {
				return err
			}
			a.printf("installed %s\n", target)
			return nil
		},
	}
	f := cmd.Flags()
	f.BoolVar(&install, "install", false, "write the skill instead of printing it")
	f.BoolVar(&project, "project", false, "install into ./.claude/skills instead of ~/.claude/skills")
	f.StringVar(&dir, "dir", "", "install into this skills directory")
	return cmd
}
