# 📚 Quiz Management System - ASP.NET Core MVC

The **Quiz Management System** is a full-featured web application developed using **ASP.NET Core MVC**, designed to manage quizzes, questions, question levels, and user access. It includes a dynamic admin dashboard, user authentication, and export functionality for all data listings.

---

## 🚀 Features

- ✅ **User Authentication**
  - Login & Registration
  - Sign-out functionality

- 📋 **CRUD Operations**
  - Users
  - Questions with options (A–D) and correct answer
  - Quizzes
  - Question Levels
  - Quiz-wise Question Mapping

- 📈 **Dashboard**
  - Dynamic count of records for each module (Users, Questions, Quizzes, etc.)

- 📤 **Export to Excel**
  - Export full data lists of any module to Excel files

- 🧭 **Responsive Sidebar UI**
  - Built with Bootstrap 5 (Admin theme)
  - Sidebar with collapsible menus for each section

---

## 🛠 Tech Stack

| Layer        | Technology            |
|--------------|------------------------|
| Frontend     | Razor Views + Bootstrap 5 |
| Backend      | ASP.NET Core MVC       |
| Database     | SQL Server             |
| ORM          | Entity Framework Core  |
| Export       | EPPlus / ClosedXML (Excel Export) |
| Auth System  | Identity-like custom login system with role flags |

---

## 🗂 Database Schema (Overview)

- **MST_User** – user accounts and roles  
- **MST_Question** – questions with options and correct answer  
- **MST_Quiz** – quiz info (name, date, total questions)  
- **MST_QuestionLevel** – difficulty levels  
- **MST_QuizWiseQuestions** – mapping of questions to quizzes  

---
