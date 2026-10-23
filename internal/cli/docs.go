package cli

import (
	_ "embed"

	"github.com/spf13/cobra"
)

//go:embed agent-docs.md
var agentDocs string

func (a *app) docsCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "agent-docs",
		Short: "Print a usage guide written for AI agents",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, args []string) error {
			a.printf("%s", agentDocs)
			return nil
		},
	}
}
