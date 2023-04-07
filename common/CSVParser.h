#ifndef ARIA_CSV_H
#define ARIA_CSV_H

#include <fstream>
#include <memory>
#include <stdexcept>
#include <string>
#include <vector>

namespace aria {
  namespace csv {
    enum class Term : wchar_t { CRLF = 65534 };
    enum class FieldType { DATA, ROW_END, CSV_END };
    typedef std::vector<std::vector<std::wstring>> CSV;

    // Checking for '\n', '\r', and '\r\n' by default
    inline bool operator==(const wchar_t c, const Term t) {
      switch (t) {
        case Term::CRLF:
          return c == L'\r' || c == L'\n';
        default:
          return static_cast<wchar_t>(t) == c;
      }
    }

    inline bool operator!=(const wchar_t c, const Term t) {
      return !(c == t);
    }

    // Wraps returned fields so we can also indicate
    // that we hit row endings or the end of the csv itself
    struct Field {
      explicit Field(FieldType t): type(t), data(nullptr) {}
      explicit Field(const std::wstring& str): type(FieldType::DATA), data(&str) {}

      FieldType type;
      const std::wstring *data;
    };

    // Reads and parses lines from a csv file
    class CsvParser {
    private:
      // CSV state for state machine
      enum class State {
        START_OF_FIELD,
        IN_FIELD,
        IN_QUOTED_FIELD,
        IN_ESCAPED_QUOTE,
        END_OF_ROW,
        EMPTY
      };
      State m_state;

      // Configurable attributes
      wchar_t m_quote;
      wchar_t m_delimiter;
      Term m_terminator;
      std::wistream& m_input;

      // Buffer capacities
      static const int FIELDBUF_CAP = 1024;
      static const int INPUTBUF_CAP = 1024 * 128;

      // Buffers
      std::wstring m_fieldbuf;
      wchar_t *m_inputbuf;

      // Misc
      bool m_eof;
      size_t m_cursor;
      size_t m_inputbuf_size;
      std::streamoff m_scanposition;
    public:
      // Creates the CSV parser which by default, splits on commas,
      // uses quotes to escape, and handles CSV files that end in either
      // '\r', '\n', or '\r\n'.
      explicit CsvParser(std::wistream& input):
        m_input(input),
        m_state(State::START_OF_FIELD),
        m_quote('"'),
        m_delimiter(','),
        m_terminator(Term::CRLF),
        m_eof(false),
        m_cursor(INPUTBUF_CAP),
        m_inputbuf(nullptr),
        m_inputbuf_size(INPUTBUF_CAP),
        m_scanposition(-INPUTBUF_CAP)
      {
        m_inputbuf = new wchar_t[INPUTBUF_CAP];
        if (m_inputbuf == nullptr) {
          throw std::runtime_error("buffer allocation error");
        }
        // Reserve space upfront to improve performance
        m_fieldbuf.reserve(FIELDBUF_CAP);
        if (!m_input.good()) {
          throw std::runtime_error("Something is wrong with input stream");
        }
      }

      // Change the quote character
      CsvParser& quote(wchar_t c) {
        m_quote = c;
        return *this;
      }

      // Change the delimiter character
      CsvParser& delimiter(wchar_t c) {
        m_delimiter = c;
        return *this;
      }

      // Change the terminator character
      CsvParser& terminator(wchar_t c) {
        m_terminator = static_cast<Term>(c);
        return *this;
      }

      // The parser is in the empty state when there are
      // no more tokens left to read from the input buffer
      bool empty() {
        return m_state == State::EMPTY;
      }

      // Not the actual position in the stream (its buffered) just the
      // position up to last availiable token
      std::streamoff position() const
      {
          return m_scanposition + static_cast<std::streamoff>(m_cursor);
      }

      // Reads a single field from the CSV
      Field next_field() {
        if (empty()) {
          return Field(FieldType::CSV_END);
        }
        m_fieldbuf.clear();

        // This loop runs until either the parser has
        // read a full field or until there's no tokens left to read
        for (;;) {
          wchar_t *maybe_token = top_token();

          // If we're out of tokens to read return whatever's left in the
          // field and row buffers. If there's nothing left, return null.
          if (!maybe_token) {
            m_state = State::EMPTY;
            return !m_fieldbuf.empty() ? Field(m_fieldbuf) : Field(FieldType::CSV_END);
          }

          // Parsing the CSV is done using a finite state machine
          wchar_t c = *maybe_token;
          switch (m_state) {
            case State::START_OF_FIELD:
              m_cursor++;
              if (c == m_terminator) {
                handle_crlf(c);
                return Field(FieldType::ROW_END);
              }

              if (c == m_quote) {
                m_state = State::IN_QUOTED_FIELD;
              } else if (c == m_delimiter) {
                return Field(m_fieldbuf);
              } else {
                m_state = State::IN_FIELD;
                m_fieldbuf += c;
              }

              break;

            case State::IN_FIELD:
              m_cursor++;
              if (c == m_terminator) {
                handle_crlf(c);
                m_state = State::END_OF_ROW;
                return Field(m_fieldbuf);
              }

              if (c == m_delimiter) {
                m_state = State::START_OF_FIELD;
                return Field(m_fieldbuf);
              } else {
                m_fieldbuf += c;
              }

              break;

            case State::IN_QUOTED_FIELD:
              m_cursor++;
              if (c == m_quote) {
                m_state = State::IN_ESCAPED_QUOTE;
              } else {
                m_fieldbuf += c;
              }

              break;

            case State::IN_ESCAPED_QUOTE:
              m_cursor++;
              if (c == m_terminator) {
                handle_crlf(c);
                m_state = State::END_OF_ROW;
                return Field(m_fieldbuf);
              }

              if (c == m_quote) {
                m_state = State::IN_QUOTED_FIELD;
                m_fieldbuf += c;
              } else if (c == m_delimiter) {
                m_state = State::START_OF_FIELD;
                return Field(m_fieldbuf);
              } else {
                m_state = State::IN_FIELD;
                m_fieldbuf += c;
              }

              break;

            case State::END_OF_ROW:
              m_state = State::START_OF_FIELD;
              return Field(FieldType::ROW_END);

            case State::EMPTY:
              throw std::logic_error("You goofed");
          }
        }
      }
    private:
      // When the parser hits the end of a line it needs
      // to check the special case of '\r\n' as a terminator.
      // If it finds that the previous token was a '\r', and
      // the next token will be a '\n', it skips the '\n'.
      void handle_crlf(const wchar_t c) {
        if (m_terminator != Term::CRLF || c != L'\r') {
          return;
        }

        wchar_t *token = top_token();
        if (token && *token == L'\n') {
          m_cursor++;
        }
      }

      // Pulls the next token from the input buffer, but does not move
      // the cursor forward. If the stream is empty and the input buffer
      // is also empty return a nullptr.
      wchar_t* top_token() {
        // Return null if there's nothing left to read
        if (m_eof && m_cursor == m_inputbuf_size) {
          return nullptr;
        }

        // Refill the input buffer if it's been fully read
        if (m_cursor == m_inputbuf_size) {
          m_scanposition += static_cast<std::streamoff>(m_cursor);
          m_cursor = 0;
          m_input.read(m_inputbuf, INPUTBUF_CAP);

          // Indicate we hit end of file, and resize
          // input buffer to show that it's not at full capacity
          if (m_input.eof()) {
            m_eof = true;
            m_inputbuf_size = (size_t)m_input.gcount();

            // Return null if there's nothing left to read
            if (m_inputbuf_size == 0) {
              return nullptr;
            }
          }
        }

        return &m_inputbuf[m_cursor];
      }
    public:
      // Iterator implementation for the CSV parser, which reads
      // from the CSV row by row in the form of a vector of strings
      class iterator {
      public:
       typedef std::ptrdiff_t difference_type;
       typedef std::vector<std::wstring> value_type;
       typedef const std::vector<std::wstring>* pointer;
       typedef const std::vector<std::wstring>& reference;
       typedef std::input_iterator_tag iterator_category;

        explicit iterator(CsvParser *p, bool end = false):
          m_parser(p),
          m_current_row(-1)
        {
          if (!end) {
            m_row.reserve(50);
            m_current_row = 0;
            next();
          }
        }

        iterator& operator++() {
          next();
          return *this;
        }

        iterator operator++(int) {
          iterator i = (*this);
          ++(*this);
          return i;
        }

        bool operator==(const iterator& other) const {
          return m_current_row == other.m_current_row
            && m_row.size() == other.m_row.size();
        }

        bool operator!=(const iterator& other) const {
          return !(*this == other);
        }

        reference operator*() const {
          return m_row;
        }

        pointer operator->() const {
          return &m_row;
        }
      private:
        value_type m_row;
        CsvParser *m_parser;
        int m_current_row;

        void next() {
          value_type::size_type num_fields = 0;
          for (;;) {
            auto field = m_parser->next_field();
            switch (field.type) {
              case FieldType::CSV_END:
                if (num_fields < m_row.size()) {
                  m_row.resize(num_fields);
                }
                m_current_row = -1;
                return;
              case FieldType::ROW_END:
                if (num_fields < m_row.size()) {
                  m_row.resize(num_fields);
                }
                m_current_row++;
                return;
              case FieldType::DATA:
                if (num_fields < m_row.size()) {
                  m_row[num_fields] = std::move(*field.data);
                } else {
                  m_row.push_back(std::move(*field.data));
                }
                num_fields++;
            }
          }
        }
      };

      iterator begin() { return iterator(this); };
      iterator end() { return iterator(this, true); };
    };
  }
}
#endif
