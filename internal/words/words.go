// Package words splits a command line into words the way a POSIX shell
// does, and quotes words so that they split back the same. Batch scripts
// (sheets do) and the styling sheets records in a workbook use it.
package words

import (
	"fmt"
	"strings"
)

// Word is one word of a line: its text, and how it was written.
type Word struct {
	Text string
	Raw  string // as written, quotes included
}

// Split breaks a line into words: blanks separate them; '...' is taken
// literally; "..." too, except that \" \\ \$ and \` are escapes; a
// backslash outside quotes takes the next character literally. A # that
// starts a word starts a comment.
func Split(line string) ([]Word, error) {
	var (
		out   []Word
		b     strings.Builder
		in    bool // inside a word
		start int
	)
	end := func(i int) {
		if in {
			out = append(out, Word{Text: b.String(), Raw: line[start:i]})
			b.Reset()
			in = false
		}
	}
	for i := 0; i < len(line); i++ {
		c := line[i]
		switch {
		case c == ' ' || c == '\t' || c == '\n' || c == '\r':
			end(i)
		case c == '#' && !in:
			return out, nil
		case c == '\'':
			if !in {
				in, start = true, i
			}
			j := strings.IndexByte(line[i+1:], '\'')
			if j < 0 {
				return nil, fmt.Errorf("unclosed ' in %s", line)
			}
			b.WriteString(line[i+1 : i+1+j])
			i += j + 1
		case c == '"':
			if !in {
				in, start = true, i
			}
			i++
			for ; i < len(line) && line[i] != '"'; i++ {
				if line[i] == '\\' && i+1 < len(line) && strings.IndexByte("\"\\$`", line[i+1]) >= 0 {
					i++
				}
				b.WriteByte(line[i])
			}
			if i >= len(line) {
				return nil, fmt.Errorf("unclosed \" in %s", line)
			}
		case c == '\\':
			if !in {
				in, start = true, i
			}
			if i+1 < len(line) {
				i++
				b.WriteByte(line[i])
			}
		default:
			if !in {
				in, start = true, i
			}
			b.WriteByte(c)
		}
	}
	end(len(line))
	return out, nil
}

// Texts splits a line and returns the words' text.
func Texts(line string) ([]string, error) {
	ws, err := Split(line)
	if err != nil {
		return nil, err
	}
	out := make([]string, len(ws))
	for i, w := range ws {
		out[i] = w.Text
	}
	return out, nil
}

// Quote returns s as one word: unchanged when it is plain, single-quoted
// otherwise.
func Quote(s string) string {
	safe := s != ""
	for i, r := range s {
		// # starts a comment only at the start of a word.
		if !(r >= 'a' && r <= 'z' || r >= 'A' && r <= 'Z' || r >= '0' && r <= '9' || strings.ContainsRune("-_./:@%+=,", r) || r == '#' && i > 0) {
			safe = false
			break
		}
	}
	if safe {
		return s
	}
	return "'" + strings.ReplaceAll(s, "'", `'\''`) + "'"
}

// Join quotes each word and joins them with spaces.
func Join(ws []string) string {
	q := make([]string, len(ws))
	for i, w := range ws {
		q[i] = Quote(w)
	}
	return strings.Join(q, " ")
}
