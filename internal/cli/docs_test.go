package cli

import (
	"bytes"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"testing"
	"unicode/utf8"
)

// Every section of the guide is a topic, every topic a section, and the
// skill's instructions point at each.
func TestDocTopics(t *testing.T) {
	sections := docSections()
	named := map[string]bool{}
	for _, tp := range docTopics {
		if _, ok := sections[tp.heading]; !ok {
			t.Errorf("topic %s: no section %q in agent-docs.md", tp.name, tp.heading)
		}
		named[tp.heading] = true
		if !strings.Contains(skillBodySource, "[["+tp.name+"]]") {
			t.Errorf("skill-body.md does not point at topic %s", tp.name)
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
	for _, body := range []string{skillFiles()["SKILL.md"], agentIndex()} {
		if strings.Contains(body, "[[") {
			t.Errorf("a topic link was left unfilled: %s", body[strings.Index(body, "[["):][:20])
		}
	}
	if _, err := docTopic("nope"); err == nil || !strings.Contains(err.Error(), "reading, query") {
		t.Errorf("unknown topic: %v", err)
	}
}

// The skill follows the Agent Skills specification (agentskills.io), as
// its reference validator (skills-ref) checks it, and the advice that goes
// with it: instructions under 500 lines and 5,000 tokens, references one
// level deep.
func TestSkillSpec(t *testing.T) {
	files := skillFiles()
	content := files["SKILL.md"]
	// skills-ref splits at the first two "---" of the file.
	parts := strings.SplitN(content, "---", 3)
	if !strings.HasPrefix(content, "---\n") || len(parts) != 3 || "---"+parts[1]+"---\n\n" != skillFront {
		t.Fatalf("frontmatter not delimited as the specification reads it")
	}
	fields, metadata := map[string]string{}, map[string]string{}
	var inMetadata bool
	for _, line := range strings.Split(strings.Trim(parts[1], "\n"), "\n") {
		if strings.HasPrefix(line, "  ") && inMetadata {
			k, v, ok := strings.Cut(strings.TrimSpace(line), ": ")
			if !ok {
				t.Errorf("metadata line %q is not key: value", line)
			}
			metadata[k] = v
			continue
		}
		k, v, _ := strings.Cut(line, ":")
		v = strings.TrimSpace(v)
		inMetadata = k == "metadata" && v == ""
		if _, dup := fields[k]; dup {
			t.Errorf("field %s twice", k)
		}
		fields[k] = v
	}
	allowed := map[string]bool{"name": true, "description": true, "license": true, "compatibility": true, "metadata": true, "allowed-tools": true}
	for k, v := range fields {
		if !allowed[k] {
			t.Errorf("field %q is not in the specification; clients refuse it", k)
		}
		// Plain YAML scalars: nothing that would read as YAML syntax.
		if strings.Contains(v, ": ") || strings.Contains(v, " #") || strings.ContainsAny(v[:min(len(v), 1)], "-?:,[]{}#&*!|>'\"%@`") && k != "metadata" {
			t.Errorf("field %s is not a plain one-line YAML value: %q", k, v)
		}
	}
	if !regexp.MustCompile(`^[a-z0-9]+(-[a-z0-9]+)*$`).MatchString(fields["name"]) || len(fields["name"]) > 64 || fields["name"] != skillName {
		t.Errorf("name %q: 1-64 lowercase letters, digits and single hyphens, the folder's name", fields["name"])
	}
	if n := utf8.RuneCountInString(fields["description"]); n == 0 || n > 1024 {
		t.Errorf("description has %d characters; 1 to 1024", n)
	}
	if n := utf8.RuneCountInString(fields["compatibility"]); n > 500 {
		t.Errorf("compatibility has %d characters; at most 500", n)
	}
	if len(metadata) == 0 {
		t.Error("metadata lost its keys")
	}
	body := parts[2]
	if n := strings.Count(body, "\n"); n >= 500 {
		t.Errorf("SKILL.md's instructions run to %d lines; keep them under 500", n)
	}
	// About four characters a token: well under the 5,000 advised.
	if n := len(body) / 4; n > 2500 {
		t.Errorf("SKILL.md's instructions are about %d tokens; keep them lean", n)
	}
	// Every link points at a file of the skill, one level down.
	for _, m := range regexp.MustCompile(`\]\(([^)]+)\)`).FindAllStringSubmatch(body, -1) {
		if _, ok := files[m[1]]; !ok || strings.Count(m[1], "/") != 1 {
			t.Errorf("link to %q, which is not a reference of the skill", m[1])
		}
	}
	for name, text := range files {
		if strings.HasPrefix(name, "references/") && !strings.HasPrefix(text, "# ") {
			t.Errorf("%s does not start with its heading", name)
		}
		// A command in backticks stays on one line, to be copied whole.
		for i, line := range strings.Split(text, "\n") {
			if !strings.HasPrefix(line, "    ") && strings.Count(line, "`")%2 == 1 {
				t.Errorf("%s:%d: a code span breaks across lines: %s", name, i+1, line)
			}
		}
	}
}

// skills/sheets in the repository is what sheets skill --install writes, so
// the skill can be installed straight from the repository.
func TestSkillFolderInRepo(t *testing.T) {
	root := filepath.Join("..", "..", "skills", skillName)
	want := skillFiles()
	got := map[string]string{}
	err := filepath.WalkDir(root, func(path string, d fs.DirEntry, err error) error {
		if err != nil || d.IsDir() {
			return err
		}
		b, err := os.ReadFile(path)
		rel, _ := filepath.Rel(root, path)
		got[filepath.ToSlash(rel)] = string(b)
		return err
	})
	if err != nil {
		t.Fatal(err)
	}
	var stale []string
	for name, text := range want {
		if got[name] != text {
			stale = append(stale, name)
		}
	}
	for name := range got {
		if _, ok := want[name]; !ok {
			stale = append(stale, name+" (not the skill's)")
		}
	}
	sort.Strings(stale)
	if len(stale) > 0 {
		t.Errorf("skills/sheets is out of date (%s); run: go run ./cmd/sheets skill --install --dir skills", strings.Join(stale, ", "))
	}
}

func TestInstallSkill(t *testing.T) {
	dir := t.TempDir()
	root := filepath.Join(dir, skillName)
	// An earlier install's reference of a topic since gone, and a file of
	// the user's own, which stays.
	for name, text := range map[string]string{"references/gone.md": "# Gone\n", "notes.txt": "mine\n"} {
		path := filepath.Join(root, filepath.FromSlash(name))
		if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(path, []byte(text), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	var out, errb bytes.Buffer
	if code := Main([]string{"skill", "--install", "--dir", dir}, nil, &out, &errb); code != 0 {
		t.Fatalf("exit %d: %s", code, errb.String())
	}
	if !strings.Contains(out.String(), "installed "+root+" (SKILL.md and 17 references)") {
		t.Errorf("output: %s", out.String())
	}
	for name, text := range skillFiles() {
		b, err := os.ReadFile(filepath.Join(root, filepath.FromSlash(name)))
		if err != nil || string(b) != text {
			t.Errorf("%s not installed as generated: %v", name, err)
		}
	}
	if _, err := os.Stat(filepath.Join(root, "references", "gone.md")); err == nil {
		t.Error("a stale reference stayed")
	}
	if _, err := os.Stat(filepath.Join(root, "notes.txt")); err != nil {
		t.Error("a file of the user's was removed")
	}
	for _, args := range [][]string{{"skill", "--agents"}, {"skill", "--install", "--dir", dir, "--project"}} {
		out.Reset()
		errb.Reset()
		if code := Main(args, nil, &out, &errb); code == 0 {
			t.Errorf("%v: accepted", args)
		}
	}
}

// In the repository's documents too, a command in backticks stays on one
// line, so it can be copied whole.
func TestDocsCodeSpans(t *testing.T) {
	for _, name := range []string{"README.md", "docs/deferred.md", "evals/README.md"} {
		b, err := os.ReadFile(filepath.Join("..", "..", filepath.FromSlash(name)))
		if err != nil {
			t.Fatal(err)
		}
		fenced := false
		for i, line := range strings.Split(string(b), "\n") {
			if strings.HasPrefix(strings.TrimSpace(line), "```") {
				fenced = !fenced
				continue
			}
			if !fenced && !strings.HasPrefix(line, "    ") && strings.Count(line, "`")%2 == 1 {
				t.Errorf("%s:%d: a code span breaks across lines: %s", name, i+1, line)
			}
		}
	}
}
