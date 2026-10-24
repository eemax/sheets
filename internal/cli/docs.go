package cli

import (
	_ "embed"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

// agentDocs is the whole guide; skillBodySource the instructions of the
// skill (see skill.go), which point at the guide's topics: [[topic]].
var (
	//go:embed agent-docs.md
	agentDocs string
	//go:embed skill-body.md
	skillBodySource string
)

// agentIndex is the skill's instructions for an agent without the skill,
// with each topic as the command that prints it.
func agentIndex() string {
	return skillBody(func(t string) string { return "`sheets agent-docs " + t + "`" }) +
		"\n`sheets agent-docs all` prints the whole guide.\n"
}

// docTopics name the guide's sections, by their headings.
var docTopics = []struct {
	name    string
	aliases []string
	heading string
}{
	{"reading", []string{"read", "info", "peek", "find"}, "Reading"},
	{"query", []string{"sql"}, "SQL queries"},
	{"addresses", []string{"references", "refs"}, "References"},
	{"writing", []string{"write", "set"}, "Writing"},
	{"do", []string{"batch"}, "Many changes at once: do"},
	{"structure", []string{"rows", "cols", "columns", "sheet", "sheets"}, "Rows, columns and sheets"},
	{"ranges", []string{"clear", "copy", "move", "sort", "replace"}, "Moving data: clear, copy, move, sort, replace"},
	{"notes", []string{"note", "link", "links"}, "Notes and links"},
	{"csv", []string{"convert"}, "CSV files"},
	{"check", []string{"lint", "diff"}, "Checking: lint and diff"},
	{"charts", []string{"chart"}, "Charts"},
	{"render", nil, "Looking at a sheet: render"},
	{"safety", []string{"undo", "dry-run"}, "Safety: dry runs and undo"},
	{"fmt", []string{"format", "formatting"}, "Formatting"},
	{"style", []string{"cond", "validate", "view"}, "Styling on request: style, cond, validate, view"},
	{"large", []string{"performance"}, "Large files"},
	{"exit", []string{"exit-status"}, "Exit status"},
}

// docSections splits the guide into its sections, by heading.
func docSections() map[string]string {
	out := map[string]string{}
	for _, part := range strings.Split(agentDocs, "\n## ")[1:] {
		heading, _, _ := strings.Cut(part, "\n")
		out[heading] = "## " + strings.TrimRight(part, "\n") + "\n"
	}
	return out
}

// docTopic is the section a topic names.
func docTopic(name string) (string, error) {
	name = strings.ToLower(strings.TrimSpace(name))
	for _, t := range docTopics {
		if name == t.name || contains(t.aliases, name) {
			if s, ok := docSections()[t.heading]; ok {
				return s, nil
			}
		}
	}
	var names []string
	for _, t := range docTopics {
		names = append(names, t.name)
	}
	return "", fmt.Errorf("no topic %q: give one of %s, or all", name, strings.Join(names, ", "))
}

func contains(list []string, s string) bool {
	for _, x := range list {
		if x == s {
			return true
		}
	}
	return false
}

func (a *app) docsCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "agent-docs [TOPIC|all]...",
		Short: "Print the guide written for AI agents: an index, topics, or all of it",
		Long: `Print the guide for AI agents. With no topic, the instructions of the sheets
skill (see sheets skill): the commands used most, the gotchas, and the
topics; with topics, those sections of the guide; with all, the whole
guide.

  sheets agent-docs
  sheets agent-docs style ranges
  sheets agent-docs all`,
		RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				a.printf("%s", agentIndex())
				return nil
			}
			var parts []string
			for _, arg := range args {
				if strings.EqualFold(arg, "all") {
					a.printf("%s", agentDocs)
					return nil
				}
				s, err := docTopic(arg)
				if err != nil {
					return err
				}
				if !contains(parts, s) {
					parts = append(parts, s)
				}
			}
			a.printf("%s", strings.Join(parts, "\n"))
			return nil
		},
	}
}
