package outlook

import (
	"bytes"
	"encoding/json"
	"errors"
	"net/http"
	"time"

	"google.golang.org/appengine/log"
	"google.golang.org/appengine/urlfetch"

	"github.com/news-ai/tabulae/models"

	"golang.org/x/net/context"
)

type EmailAddress struct {
	EmailAddress struct {
		Address string `json:"Address"`
	} `json:"EmailAddress"`
}

type EmailRequest struct {
	Message struct {
		Subject string `json:"Subject"`
		Body    struct {
			ContentType string `json:"ContentType"`
			Content     string `json:"Content"`
		} `json:"Body"`
		ToRecipients []EmailAddress `json:"ToRecipients"`
	} `json:"Message"`
}

type EmailRequestAttachment struct {
	Message struct {
		Subject string `json:"Subject"`
		Body    struct {
			ContentType string `json:"ContentType"`
			Content     string `json:"Content"`
		} `json:"Body"`
		ToRecipients []struct {
			EmailAddress struct {
				Address string `json:"Address"`
			} `json:"EmailAddress"`
		} `json:"ToRecipients"`
		Attachments []struct {
			OdataType    string `json:"@odata.type"`
			Name         string `json:"Name"`
			ContentBytes string `json:"ContentBytes"`
		} `json:"Attachments"`
	} `json:"Message"`
}

func (o *Outlook) SendEmail(c context.Context, from string, to string, subject string, body string, email models.Email) error {
	if len(o.AccessToken) > 0 {
		contextWithTimeout, _ := context.WithTimeout(c, time.Second*15)
		client := urlfetch.Client(contextWithTimeout)

		toEmail := EmailAddress{}
		toEmail.EmailAddress.Address = to

		var message EmailRequest

		message.Message.Subject = subject
		message.Message.ToRecipients = append(message.Message.ToRecipients, toEmail)

		message.Message.Body.ContentType = "HTML"
		message.Message.Body.Content = body

		messageJson, err := json.Marshal(message)
		if err != nil {
			log.Errorf(c, "%v", err)
			return err
		}

		messageQuery := bytes.NewReader(messageJson)

		URL := BASEURL + "api/v2.0/me/sendmail"
		req, _ := http.NewRequest("POST", URL, messageQuery)

		req.Header.Add("Authorization", "Bearer "+o.AccessToken)
		req.Header.Add("Content-Type", "application/json")

		response, err := client.Do(req)
		if err != nil {
			log.Errorf(c, "%v", err)
			return err
		}

		if response.StatusCode == 200 || response.StatusCode == 202 {
			return nil
		}

		return errors.New("Email could not be sent")
	}

	return errors.New("No access token supplied")
}

func (o *Outlook) SendEmailWithAttachments(r *http.Request, c context.Context, from string, to string, subject string, body string, email models.Email, files []models.File) error {
	return errors.New("No access token supplied")
}
