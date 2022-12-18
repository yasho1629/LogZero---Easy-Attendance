const express = require("express");
const bodyParser = require("body-parser");
const date = require(__dirname + "/date.js");
const mongoose = require("mongoose");
const exceljS = require("exceljs");
const exportUsersToExcel = require('./exportService');
const exportMarksToExcel = require('./exportService');
mongoose.connect("mongodb://localhost:27017/DBMS", {
  useNewUrlParser: true,
  useUnifiedTopology: true
});
mongoose.set('useFindAndModify', false);

const app = express();

app.set('view engine', 'ejs');

app.use(bodyParser.urlencoded({
  extended: true
}));
app.use(express.static("public"));

const facultySchema = new mongoose.Schema({
  faculty_id: Number,
  username: String,
  password: String,
  course: String,
  phone_num: Number
});

const Faculty = mongoose.model("Faculty", facultySchema);

const faculty1 = new Faculty({
  faculty_id: 1,
  username: "Ranjana Badre",
  password: "1234",
  course: "DBMS",
  phone_num: 1234567890
});

const faculty2 = new Faculty({
  faculty_id: 2,
  username: "Sanjay Ghodke",
  password: "1234",
  course: "COA",
  phone_num: 1234567891
});

const faculty3 = new Faculty({
  faculty_id: 3,
  username: "Pramod Ganjewar",
  password: "1234",
  course: "ADS",
  phone_num: 1234567892
});

const faculty4 = new Faculty({
  faculty_id: 4,
  username: "Preeti Shinde",
  password: "1234",
  course: "AM",
  phone_num: 1234567893
});

const faculty5 = new Faculty({
  faculty_id: 5,
  username: "Nilesh Bhandare",
  password: "1234",
  course: "PS",
  phone_num: 1234567894
});

const studentSchema = new mongoose.Schema({
  roll_number: Number,
  name: String,
  prn_number: Number
});

const Student = mongoose.model("Student", studentSchema);

const student1 = new Student({
  roll_number: 1,
  name: "Aaditya Barve",
  prn_number: 120190001
});

const student2 = new Student({
  roll_number: 2,
  name: "Arya Kashyap",
  prn_number: 120190002
});

const student3 = new Student({
  roll_number: 3,
  name: "Sarang Wadode",
  prn_number: 120190003
});

const student4 = new Student({
  roll_number: 4,
  name: "Manoranjan Jena",
  prn_number: 120190004
});

const student5 = new Student({
  roll_number: 5,
  name: "Pranav Hatwar",
  prn_number: 120190005
});

const student6 = new Student({
  roll_number: 6,
  name: "Manish Choudhary",
  prn_number: 120190006
});

const student7 = new Student({
  roll_number: 7,
  name: "Shreya Gaikwad",
  prn_number: 120190007
});

const student8 = new Student({
  roll_number: 8,
  name: "Aditya Kumar",
  prn_number: 120190008
});

const student9 = new Student({
  roll_number: 9,
  name: "Sanika Pareek",
  prn_number: 120190009
});

const student10 = new Student({
  roll_number: 10,
  name: "Niranjan Girhe",
  prn_number: 120190010
});

const student11 = new Student({
  roll_number: 11,
  name: "Mayuresh Kumar",
  prn_number: 120190011
});

const student12 = new Student({
  roll_number: 12,
  name: "Tejas Kalje",
  prn_number: 120190012
});

const student13 = new Student({
  roll_number: 13,
  name: "Shwetali Desai",
  prn_number: 120190013
});

const student14 = new Student({
  roll_number: 14,
  name: "Aventika Khemani",
  prn_number: 120190014
});

const student15 = new Student({
  roll_number: 15,
  name: "Chirayu Shah",
  prn_number: 120190015
});

const student16 = new Student({
  roll_number: 16,
  name: "Atharva Khonde",
  prn_number: 120190016
});

const student17 = new Student({
  roll_number: 17,
  name: "Rahul Pradhan",
  prn_number: 120190017
});

const student18 = new Student({
  roll_number: 18,
  name: "Aditi Borkar",
  prn_number: 120190018
});

const student19 = new Student({
  roll_number: 19,
  name: "Pratik Kale",
  prn_number: 120190019
});

const student20 = new Student({
  roll_number: 20,
  name: "Ranjeet Bhosale",
  prn_number: 120190021
});

const attendanceSchema = new mongoose.Schema({
  faculty_id: Number,
  roll_number: Number,
  attendance_status: String,
  date: String,
  course: String
});

const Attendance = mongoose.model("Attendance", attendanceSchema);

const marksSchema = new mongoose.Schema({
  faculty_id: Number,
  roll_number: Number,
  marks: Number,
  course: String
});

const Marks = mongoose.model("Mark", marksSchema);

app.get("/", function(req,res){
  res.render("index",{});
});

app.get("/login", function(req,res){
  Faculty.find(function(err, faculties){
    if(err)
      console.log(err);
    else if(faculties.length === 0)
    {
      Faculty.insertMany([faculty1, faculty2, faculty3, faculty4, faculty5], function(err){
        if(err)
          console.log(err);
        else
          console.log("Faculty Details successfully inserted");
      });
      res.redirect("/login");
    }
    else
      res.render("login",{errorMessage: "No Error"});
  });
});

app.get("/register",function(req,res){
  res.render("register");
});

var username="";
var facultyId=0;
var course="";

app.post("/login", function(req,res){
  const enteredUsername = req.body.username;
  const enteredPassword = req.body.password;
  Faculty.findOne({username: enteredUsername, password: enteredPassword},function(err, faculty){
    if(err)
      console.log(err);
    else if(faculty)
    {
      username = enteredUsername;
      facultyId = faculty.faculty_id;
      course = faculty.course;
      console.log(course);
      res.redirect("/home");
    }
    else
      res.render("login",{errorMessage: "Error"});
  })
});

var fac_id=6;

app.post("/register",function(req,res){
  const enteredUsername = req.body.username;
  const enteredPassword = req.body.password;
  const enteredCourse = req.body.course;
  const enteredPhone = Number(req.body.phone);
  const faculty = new Faculty({faculty_id: fac_id, username: enteredUsername, password: enteredPassword, course: enteredCourse, phone_num: enteredPhone});
  faculty.save();
  username = enteredUsername;
  facultyId = fac_id;
  fac_id+=1;
  course=enteredCourse;
  console.log("Faculty Details successfully inserted");
  res.redirect("/home");
});

app.get("/home", function(req,res){
  res.render("home", {username: username});
});

var day;

app.get("/attendance", function(req,res){
  Student.find(function(err,students){
    if(err)
      console.log(err);
    else if(students.length===0)
    {
      Student.insertMany([student1, student2, student3, student4, student5, student6, student7, student8, student9, student10, student11, student12, student13, student14, student15, student16, student17, student18, student19, student20], function(err){
        if(err)
          console.log(err);
        else
          console.log("Student Details inserted successfully");
      });
      res.redirect("/attendance");
    }
    else
    {
      day=date.getDate();
      Attendance.find({course: course, date: day}, function(err,attendance_records){
        if(err)
          console.log(err);
        else if(attendance_records.length===0)
        {
          for(var i=0;i<students.length;i++)
          {
            const student_attendance = new Attendance({
              faculty_id: facultyId,
              roll_number: students[i].roll_number,
              attendance_status: "Present",
              date: day,
              course: course
            });
            student_attendance.save();
            console.log("Inserted attendance");
          }
          res.redirect("/attendance")
        }
        else
          res.render("attendance", {studentList: students, listTitle: day, attendanceList: attendance_records});
      });
    }
  });
});

var length=20;

app.post("/selectDate",function(req,res){
  day=req.body.date;
  Student.find(function(err,students){
    if(err)
      console.log(err);
    else if(students.length===0)
    {
      Student.insertMany([student1, student2, student3, student4, student5, student6, student7, student8, student9, student10, student11, student12, student13, student14, student15, student16, student17, student18, student19, student20], function(err){
        if(err)
          console.log(err);
        else
          console.log("Student Details inserted successfully");
      });
      res.redirect("/attendance");
    }
    else
    {
      Attendance.find({course: course, date: day}, function(err,attendance_records){
        if(err)
          console.log(err);
        else if(attendance_records.length===0)
        {
          for(var i=0;i<students.length;i++)
          {
            const student_attendance = new Attendance({
              faculty_id: facultyId,
              roll_number: students[i].roll_number,
              attendance_status: "Present",
              date: day,
              course: course
            });
            student_attendance.save();
            console.log("Inserted Attendance");
          }
          res.render("attendance", {studentList: students, listTitle: day, attendanceList: attendance_records});
        }
        else
          res.render("attendance", {studentList: students, listTitle: day, attendanceList: attendance_records});
      });
    }
  });
});

app.post("/add",function(req,res){
  const arr=[];
  for(var i=0;i<length;i++)
  {
    const status = req.body.status[i];
    Attendance.updateOne({roll_number: i+1, course: course, date: day}, {attendance_status: status}, function(err){
      if(err)
        console.log(err);
      else
        console.log("successfully updated attendance");
    });
  }
  res.redirect("/attendance");
});

app.get("/marks", function(req,res){
  day=date.getDate();
  Student.find(function(err,students){
    if(err)
      console.log(err);
    else
    {
      Marks.find({course: course},function(err,marks_records){
        if(err)
          console.log(err);
        else if(marks_records.length===0)
        {
          for(var i=0;i<students.length;i++)
          {
              const student_marks = new Marks({
              faculty_id: facultyId,
              roll_number: students[i].roll_number,
              marks: 0,
              course: course
            });
            student_marks.save();
          }
          res.redirect("/marks");
        }
        else
          res.render("IA", {studentList: students, listTitle: day, markList: marks_records});
      });
    }
  });
});

app.post("/addMarks", function(req,res){
  for(var i=0;i<length;i++)
  {
    if(req.body.marks[i]!="")
    {
      const imarks = Number(req.body.marks[i]);
      Marks.updateOne({roll_number: i+1, course: course}, {marks: imarks}, function(err){
        if(err)
          console.log(err);
        else
          console.log("successfully updated marks");
      })
    }
  }
  res.redirect("/marks");
});

const finalAttendanceSchema = new mongoose.Schema({
  roll_number: Number,
  name: String,
  dbms: Number,
  coa: Number,
  ads: Number,
  am: Number,
  ps: Number
});

const FinalAttendance = mongoose.model("aResult", finalAttendanceSchema);

const finalMarksSchema = new mongoose.Schema({
  roll_number: Number,
  name: String,
  dbms: Number,
  coa: Number,
  ads: Number,
  am: Number,
  ps: Number
});

const FinalMarks = mongoose.model("mresult", finalMarksSchema);

/*
app.get("/exportAttendance",function(req,res)
{
  for(var i=0;i<length;i++)
  {
    const rollNum=i+1;
    Attendance.find({roll_number: rollNum}, function(err, total_records){
      if(err)
        console.log(err);
      else
      {
        Attendance.find({roll_number: rollNum, attendance_status: "Present"}, function(err, present_records){
          if(err)
            console.log(err);
          else
          {
            console.log(present_records.length*100/total_records.length);
            var student_attendance_result = new FinalAttendance({
            roll_number: rollNum,
            Attendance_Percentage: present_records.length*100/total_records.length
            });
            student_attendance_result.save();
          }
        });
      }
    });
  }
  res.redirect("/home");
});
*/
app.get("/exportAttendance",function(req,res){
  var dbms_res=0,coa_res=0,ads_res=0,am_res=0,ps_res=0;
  var name="";
  for(var i=0;i<length;i++)
  {
    const rollNum=i+1;
    Attendance.find({roll_number: rollNum, course: "DBMS"}, function(err, dtotal_records){
        Attendance.find({roll_number: rollNum, attendance_status: "Present", course: "DBMS"}, function(err, dpresent_records){
            dbms_res=dpresent_records.length*100/dtotal_records.length;
            Attendance.find({roll_number: rollNum, course: "COA"}, function(err, ctotal_records){
                Attendance.find({roll_number: rollNum, attendance_status: "Present", course: "COA"}, function(err, cpresent_records){
                    coa_res=cpresent_records.length*100/ctotal_records.length;
                    Attendance.find({roll_number: rollNum, course: "ADS"}, function(err, atotal_records){
                        Attendance.find({roll_number: rollNum, attendance_status: "Present", course: "ADS"}, function(err, apresent_records){
                            ads_res=apresent_records.length*100/atotal_records.length;
                            Attendance.find({roll_number: rollNum, course: "AM"}, function(err, amtotal_records){
                                Attendance.find({roll_number: rollNum, attendance_status: "Present", course: "AM"}, function(err, ampresent_records){
                                    am_res=ampresent_records.length*100/amtotal_records.length;
                                    Attendance.find({roll_number: rollNum, course: "PS"}, function(err, ptotal_records){
                                        Attendance.find({roll_number: rollNum, attendance_status: "Present", course: "PS"}, function(err, ppresent_records){
                                            console.log(ppresent_records.length*100/ptotal_records.length);
                                            ps_res=ppresent_records.length*100/ptotal_records.length;
                                            Student.find({roll_number: rollNum}, function(err, records){
                                                name=records[0].name;
                                                const student_attendance = new FinalAttendance({
                                                  roll_number: rollNum,
                                                  name: name,
                                                  dbms: dbms_res,
                                                  coa: coa_res,
                                                  ads: ads_res,
                                                  am: am_res,
                                                  ps: ps_res
                                                });
                                                student_attendance.save();
                                            });
                                        });
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    });
  }
  FinalAttendance.find(function(err, attendance_records){
    if(err)
      console.log(err);
    else
    {
      const workSheetColumnName = [
          "Roll",
          "Name",
          "DBMS",
          "COA",
          "ADS",
          "AM",
          "PS"
      ]
      var attendance = [];
      console.log(attendance_records);
      for(var i=0;i<attendance_records.length;i++)
      {
        var record={
          roll_number: attendance_records[i].roll_number,
          name: attendance_records[i].name,
          dbms_res: attendance_records[i].dbms,
          coa_res: attendance_records[i].coa,
          ads_res: attendance_records[i].ads,
          am_res: attendance_records[i].am,
          ps_res: attendance_records[i].ps
        }
        attendance.push(record);
      }
      console.log(attendance);
      const workSheetName = 'Attendance';
      const filePath = 'attendance.xlsx';
      exportUsersToExcel(attendance, workSheetColumnName, workSheetName, filePath);
    }
  });
  res.redirect("/home");
});

app.get("/exportMarks",function(req,res){
  var dbms_res=0,coa_res=0,ads_res=0,am_res=0,ps_res=0;
  var fname="";
  for(var i=0;i<length;i++)
  {
    const rollNum=i+1;
    /*
    Marks.findOne({roll_number: rollNum, course: "DBMS"}, function(err,drecord){
      dbms_res=drecord.marks;
      Marks.findOne({roll_number: rollNum, course: "COA"}, function(err,crecord){
        coa_res=crecord.marks;
        Marks.findOne({roll_number: rollNum, course: "ADS"}, function(err,arecord){
          ads_res=arecord.marks;
          Marks.findOne({roll_number: rollNum, course: "AM"}, function(err,amrecord){
            am_res=amrecord.marks;
            Marks.findOne({roll_number: rollNum, course: "PS"}, function(err,precord){
              ps_res=precord.marks;
              Marks.find({roll_number: rollNum}, function(err,record){
                name=record[0].name;
                const student_marks = new FinalMarks({
                  roll_number: rollNum,
                  name: name,
                  dbms: dbms_res,
                  coa: coa_res,
                  ads: ads_res,
                  am: am_res,
                  ps: ps_res
                });
                student_marks.save();
              });
            });
          });
        });
      });
    });*/
    Marks.find({roll_number: rollNum}, function(err, record){
      for(var j=0;j<record.length;j++)
      {
        if(record[j].course==="DBMS")
        {
          dbms_res=record[j].marks;
        }
        else if(record[j].course==="COA")
        {
          coa_res=record[j].marks;
        }
        else if(record[j].course==="AM")
        {
          am_res=record[j].marks;
        }
        else if(record[j].course==="PS")
        {
          ps_res=record[j].marks;
        }
        else if(record[j].course==="ADS")
        {
          ads_res=record[j].marks;
        }
      }
      Student.findOne({roll_number: rollNum}, function(err, srecord){
        fname=srecord.name;
        const student_marks = new FinalMarks({
          roll_number: rollNum,
          name: fname,
          dbms: dbms_res,
          coa: coa_res,
          ads: ads_res,
          am: am_res,
          ps: ps_res
        });
        student_marks.save();
      });
    });
  }
  FinalMarks.find(function(err, marks_records){
    if(err)
      console.log(err);
    else
    {
      const workSheetColumnName = [
          "Roll",
          "Name",
          "DBMS",
          "COA",
          "ADS",
          "AM",
          "PS"
      ]
      var marks = [];
      console.log(marks_records);
      for(var i=0;i<marks_records.length;i++)
      {
        var record={
          roll_number: marks_records[i].roll_number,
          name: marks_records[i].name,
          dbms_res: marks_records[i].dbms,
          coa_res: marks_records[i].coa,
          ads_res: marks_records[i].ads,
          am_res: marks_records[i].am,
          ps_res: marks_records[i].ps
        }
        marks.push(record);
      }
      console.log(marks);
      const workSheetName = 'Marks';
      const filePath = 'marks.xlsx';
      exportMarksToExcel(marks, workSheetColumnName, workSheetName, filePath);
    }
  });
  res.redirect("/home");
});

app.listen(3000, function() {
  console.log("Server started on port 3000");
});
