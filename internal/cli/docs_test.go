package cli

import (
	"strings"
	"testing"
)

// Every section of the guide is a topic, every topic a section, and the
// index names them all.
func TestDocTopics(t *testing.T) {
	sections := docSections()
	named := map[string]bool{}
	for _, tp := range docTopics {
		if _, ok := sections[tp.heading]; !ok {
			t.Errorf("topic %s: no section %q in agent-docs.md", tp.name, tp.heading)
		}
		named[tp.heading] = true
		if !strings.Contains(agentIndex, tp.name) {
			t.Errorf("the index does not name topic %s", tp.name)
		}
		for _, name := range append([]string{tp.name}, tp.aliases...) {
			s, err := docTopic(name)
			if err != nil || !strings.HasPrefix(s, "## "+tp.heading+"\n") {
				t.Errorf("topic %s: %q, %v", name, s[:min(len(s), 40)], err)
			}
		}
	}
	for heading := range sections {
		if !named[heading] {
			t.Errorf("section %q has no topic", heading)
		}
	}
	if _, err := docTopic("nope"); err == nil || !strings.Contains(err.Error(), "reading, query") {
		t.Errorf("unknown topic: %v", err)
	}
	// The skill loads the index into every conversation it is used in:
	// it stays short.
	if n := len(strings.Fields(agentIndex)); n > 520 {
		t.Errorf("the index has %d words; keep it near 500 (about 700 tokens)", n)
	}
}
