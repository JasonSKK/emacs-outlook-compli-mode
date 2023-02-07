;;; emacs-outlook-compli-mode.el --- outlook mode for composing and sending email (ONLY)

;; Copyright (C) 2023 Iason SK

;; Author: Iason SK <jason.skk98[at]gmail[dot]>
;; Keywords: outlook, Emacs, compliance.

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:
;; This Emacs major mode allows users to send emails through Microsoft
;; Outlook by utilising simple osa-scripting (MacOs only).  You will
;; therefore need the Outlook application.  It eliminates the need for
;; using fake email client IDs, making it ideal for those who wish to
;; maintain compliance with organisation regulations.  With this mode,
;; users can automate the process of composing and sending emails from
;; within Emacs.

;; What this mode does NOT do:
;; DOES NOT communicate with outlook/exchange server in any way.
;; DOES NOT Fetch and read email yet (todo)
;; DOES NOT Send Attachments
;; DOES NOT Support text signature
;; DOES NOT do many more

;;; Code:
(define-derived-mode outlook-compli-mode text-mode "outlook-compli"
  "Major mode for email compositions through osascripting.")

(defun outlook-compose-mail ()
  "Create a new buffer for outlook email composing and switch to it."
  (interactive)
  (switch-to-buffer (generate-new-buffer "outlook compose message"))
  (outlook-mode)
  (insert
   (concat "From:" outlook-address1 "\nTo: \nSubject: \n" )))

(defun outlook-mode-settings ()
  "Settings for `outlook-mode`."
  (setq fill-column 72) ; wrap lines at 72 characters
  (turn-on-auto-fill))

(add-hook 'outlook-mode-hook 'outlook-mode-settings)

(defun outlook-message-send ()
  "Get the text after the ':' in the first three lines of the buffer and use it as arguments for the outlook-send-message function.
Also, get the text after the 4th line and pass it as an argument as a string."
  (interactive)
  (goto-char (point-min))
  (let ((from-line (buffer-substring-no-properties (line-beginning-position) (line-end-position)))
        (to-line (progn (forward-line) (buffer-substring-no-properties (line-beginning-position) (line-end-position))))
        (subject-line (progn (forward-line) (buffer-substring-no-properties (line-beginning-position) (line-end-position)))))
    (let ((from (substring from-line (1+ (string-match ":" from-line)) (length from-line)))
          (to (substring to-line (1+ (string-match ":" to-line)) (length to-line)))
          (subject (substring subject-line (1+ (string-match ":" subject-line)) (length subject-line))))
      (forward-line)
      (let ((body (buffer-substring-no-properties (line-beginning-position) (point-max))))
        (outlook-osascript from subject body to)
        (message "Sending message from '%s' to '%s' with subject '%s' and body '%s'" from to subject body)
        ))))

(defun outlook-osascript (from subject body to)
"Contains the actual osascript command for Outlook."
(interactive)

(shell-command (format "osascript -e 'tell application \"Microsoft Outlook\"
    set theMessage to make new outgoing message with properties {sender:{name:\"YOURNAME\", address:\"%s\"}, subject:\"%s\", plain text content:\"%s\"}

    tell theMessage
        make new to recipient with properties {email address:{address:\"%s\"}}
    end tell
    send theMessage
end tell'" from subject body to)))


(define-key outlook-mode-map (kbd "C-c C-c") 'outlook-message-send)

;; set primary address
(setq outlook-address1 "youremail@example.co.uk")

(provide 'outlook-mode)
