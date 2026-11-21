(function(global) {
    "use strict";

    // ข้อมูลปริศนาทั้งหมด 90 ข้อ (6 ชั้น ชั้นละ 15 ข้อ) รองรับทั้งภาษาไทยและภาษาอังกฤษ
    const clueDatabase = [
        // ==========================================
        // ชั้น 6: จุดเริ่มต้นของการระบาด (ง่ายที่สุด)
        // ==========================================
        {
            id: "F6_C01", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "มีไฟเตือน 2 ดวงติดอยู่ แล้วมีไฟติดเพิ่มอีก 3 ดวง ขณะนี้มีไฟติดอยู่ทั้งหมดกี่ดวง",
            question_en: "There are 2 warning lights on, and 3 more lights turn on. How many lights are on in total now?",
            answer: "5",
            explanation: "นำไฟเดิม 2 ดวงบวกกับไฟที่ติดเพิ่มอีก 3 ดวง ได้ทั้งหมด 5 ดวง",
            explanation_en: "Add the initial 2 lights and the 3 additional lights to get 5 lights in total."
        },
        {
            id: "F6_C02", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "คุณพบผู้รอดชีวิตในห้องแรก 4 คน และในห้องที่สองอีก 6 คน รวมมีผู้รอดชีวิตทั้งหมดกี่คน",
            question_en: "You find 4 survivors in the first room and 6 more in the second room. How many survivors are there in total?",
            answer: "10",
            explanation: "4 บวก 6 เท่ากับ 10",
            explanation_en: "4 plus 6 equals 10."
        },
        {
            id: "F6_C03", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "เสียงประกาศเตือนภัยดังขึ้นทุกๆ 10 นาที หากครั้งแรกดังตอน 08:00 น. ครั้งที่สามจะดังเวลาใด (ตอบในรูปแบบ HH:MM)",
            question_en: "An alarm announcement sounds every 10 minutes. If the first one is at 08:00, what time will the third one sound? (Answer in HHMM format)",
            answer: "0820",
            explanation: "ครั้งแรก 08:00, ครั้งที่สอง 08:10, ครั้งที่สาม 08:20",
            explanation_en: "First at 08:00, second at 08:10, third at 08:20."
        },
        {
            id: "F6_C04", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "ตู้เก็บของมีรหัสผ่านเป็นเลขเรียงลำดับ: 2, 4, 6, 8, ... เลขตัวต่อไปคืออะไร",
            question_en: "A locker passcode follows the sequence: 2, 4, 6, 8, ... What is the next number?",
            answer: "10",
            explanation: "เป็นลำดับเลขคู่ที่เพิ่มขึ้นทีละ 2",
            explanation_en: "It is a sequence of even numbers increasing by 2."
        },
        {
            id: "F6_C05", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "มีอาหารกระป๋อง 15 กระป๋อง ถูกหยิบไปกิน 4 กระป๋อง จะเหลืออาหารกี่กระป๋อง",
            question_en: "There are 15 cans of food, and 4 cans are taken to eat. How many cans of food remain?",
            answer: "11",
            explanation: "15 ลบด้วย 4 เท่ากับ 11",
            explanation_en: "15 minus 4 equals 11."
        },
        {
            id: "F6_C06", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "รหัสปลดล็อคประตูกระจกคือผลคูณของ 3 และ 7 รหัสนั้นคืออะไร",
            question_en: "The glass door unlock code is the product of 3 and 7. What is the code?",
            answer: "21",
            explanation: "3 คูณ 7 เท่ากับ 21",
            explanation_en: "3 multiplied by 7 equals 21."
        },
        {
            id: "F6_C07", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "วันนี้คือวันที่ 5 หากการอพยพจะเริ่มขึ้นในอีก 3 วันข้างหน้า การอพยพจะตรงกับวันที่เท่าไหร่",
            question_en: "Today is the 5th. If the evacuation begins in 3 days, on what day of the month will it happen?",
            answer: "8",
            explanation: "5 บวก 3 เท่ากับ 8",
            explanation_en: "5 plus 3 equals 8."
        },
        {
            id: "F6_C08", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "ในห้องพยาบาลมีเตียง 12 เตียง มีผู้ป่วยนอนอยู่ 7 เตียง มีเตียงว่างกี่เตียง",
            question_en: "The infirmary has 12 beds, and 7 beds are occupied by patients. How many empty beds are there?",
            answer: "5",
            explanation: "12 ลบ 7 เท่ากับ 5",
            explanation_en: "12 minus 7 equals 5."
        },
        {
            id: "F6_C09", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "อุณหภูมิในห้องคือ 25 องศา หากเครื่องปรับอากาศลดอุณหภูมิลง 4 องศา อุณหภูมิใหม่จะเป็นเท่าไหร่",
            question_en: "The room temperature is 25 degrees. If the air conditioner lowers it by 4 degrees, what will the new temperature be?",
            answer: "21",
            explanation: "25 ลบ 4 เท่ากับ 21",
            explanation_en: "25 minus 4 equals 21."
        },
        {
            id: "F6_C10", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "มีขวดน้ำ 3 แพ็ค แพ็คละ 6 ขวด รวมมีขวดน้ำทั้งหมดกี่ขวด",
            question_en: "There are 3 packs of water bottles with 6 bottles in each pack. How many water bottles are there in total?",
            answer: "18",
            explanation: "3 คูณ 6 เท่ากับ 18",
            explanation_en: "3 multiplied by 6 equals 18."
        },
        {
            id: "F6_C11", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "รหัสเปิดกล่องฉุกเฉินคือตัวเลขที่หายไป: 10, 20, 30, __, 50 ตัวเลขนั้นคืออะไร",
            question_en: "The emergency box code is the missing number: 10, 20, 30, __, 50. What is that number?",
            answer: "40",
            explanation: "เป็นลำดับตัวเลขที่เพิ่มขึ้นทีละ 10",
            explanation_en: "It is a sequence increasing by 10."
        },
        {
            id: "F6_C12", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "หน้ากากอนามัย 1 กล่องมี 50 ชิ้น หากแบ่งให้คน 10 คน คนละเท่าๆ กัน จะได้คนละกี่ชิ้น",
            question_en: "One box of surgical masks contains 50 masks. If divided equally among 10 people, how many masks does each person get?",
            answer: "5",
            explanation: "50 หารด้วย 10 เท่ากับ 5",
            explanation_en: "50 divided by 10 equals 5."
        },
        {
            id: "F6_C13", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "ถังดับเพลิงมีหมายเลข 112, 113, 114 ถังถัดไปควรเป็นหมายเลขอะไร",
            question_en: "Fire extinguishers are numbered 112, 113, 114. What should the next extinguisher's number be?",
            answer: "115",
            explanation: "หมายเลขเพิ่มขึ้นทีละ 1",
            explanation_en: "The numbers increase by 1."
        },
        {
            id: "F6_C14", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "บันไดมีทั้งหมด 20 ขั้น คุณเดินลงมาแล้ว 8 ขั้น เหลืออีกกี่ขั้นจึงจะถึงชั้นล่าง",
            question_en: "A staircase has 20 steps in total. You have walked down 8 steps. How many steps remain to reach the bottom?",
            answer: "12",
            explanation: "20 ลบ 8 เท่ากับ 12",
            explanation_en: "20 minus 8 equals 12."
        },
        {
            id: "F6_C15", floor: 6, theme: "จุดเริ่มต้นของการระบาด", difficulty: "ง่าย",
            question: "หากรหัสผ่านคือเลขคี่ที่อยู่ระหว่าง 6 และ 8 รหัสนั้นคือเลขอะไร",
            question_en: "If the password is the odd number between 6 and 8, what is the password?",
            answer: "7",
            explanation: "เลขคี่ตัวเดียวที่อยู่ระหว่าง 6 และ 8 คือ 7",
            explanation_en: "The only odd number between 6 and 8 is 7."
        },

        // ==========================================
        // ชั้น 5: สำนักงานร้าง (ง่ายถึงปานกลาง)
        // ==========================================
        {
            id: "F5_C01", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "รหัสผ่านคอมพิวเตอร์คือผลรวมของรหัสพนักงานสามคน: 105, 210 และ 305 รหัสผ่านคืออะไร",
            question_en: "The computer password is the sum of three employee IDs: 105, 210, and 305. What is the password?",
            answer: "620",
            explanation: "105 + 210 + 305 = 620",
            explanation_en: "105 + 210 + 305 = 620."
        },
        {
            id: "F5_C02", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "บนโต๊ะมีแฟ้มเอกสารจัดเรียงตามปี: 2018, 2019, 2020 หากต้องหยิบแฟ้มของปีถัดไป 2 ปี จะต้องหยิบแฟ้มปีอะไร",
            question_en: "Files on the desk are arranged by year: 2018, 2019, 2020. If you need the file from 2 years after the latest one, which year do you pick?",
            answer: "2022",
            explanation: "ปีล่าสุดคือ 2020 ถัดไปอีก 2 ปีคือ 2022",
            explanation_en: "The latest year is 2020; 2 years later is 2022."
        },
        {
            id: "F5_C03", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "ล็อกเกอร์พนักงานมีรหัส 4 หลัก โดยสองหลักแรกคือ 12 และสองหลักหลังคือผลคูณของสองหลักแรก รหัสคืออะไร",
            question_en: "A staff locker has a 4-digit code. The first two digits are 12, and the last two digits are the product of the first two digits. What is the code?",
            answer: "1202",
            explanation: "1 คูณ 2 ได้ 2 (เขียนเป็น 02) รวมเป็น 1202",
            explanation_en: "1 multiplied by 2 is 2 (written as 02), making 1202."
        },
        {
            id: "F5_C04", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "พนักงาน A เข้างาน 08:30 น. พนักงาน B เข้างานสายกว่า A 45 นาที พนักงาน B เข้างานเวลาใด (ตอบ HH:MM)",
            question_en: "Employee A arrives at 08:30. Employee B arrives 45 minutes later than A. What time does Employee B arrive? (Answer in HHMM format)",
            answer: "0915",
            explanation: "08:30 บวก 45 นาที เท่ากับ 09:15",
            explanation_en: "08:30 plus 45 minutes equals 09:15."
        },
        {
            id: "F5_C05", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "กระดาษโน้ตถูกฉีกขาด มีตัวเลข 1, 4, 9, 16, __ ตัวเลขที่หายไปคืออะไร",
            question_en: "A torn sticky note shows the numbers: 1, 4, 9, 16, __. What is the missing number?",
            answer: "25",
            explanation: "เป็นลำดับเลขยกกำลังสอง: 1^2, 2^2, 3^2, 4^2, ดังนั้นตัวต่อไปคือ 5^2 = 25",
            explanation_en: "It is a sequence of squares: 1^2, 2^2, 3^2, 4^2, so the next is 5^2 = 25."
        },
        {
            id: "F5_C06", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "เครื่องปริ้นเตอร์พิมพ์งานได้ 15 แผ่นต่อนาที หากต้องการพิมพ์เอกสาร 60 แผ่น ต้องใช้เวลากี่นาที",
            question_en: "A printer prints 15 pages per minute. How many minutes does it take to print a 60-page document?",
            answer: "4",
            explanation: "60 หาร 15 เท่ากับ 4",
            explanation_en: "60 divided by 15 equals 4."
        },
        {
            id: "F5_C07", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "รหัสเปิดลิ้นชักผู้จัดการคือจำนวนโต๊ะทำงานในห้องคูณด้วย 5 หากมีโต๊ะ 14 ตัว รหัสคืออะไร",
            question_en: "The manager's drawer code is the number of desks in the room multiplied by 5. If there are 14 desks, what is the code?",
            answer: "70",
            explanation: "14 คูณ 5 เท่ากับ 70",
            explanation_en: "14 multiplied by 5 equals 70."
        },
        {
            id: "F5_C08", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "ตู้เอกสารเรียงตามตัวอักษร A=1, B=2, C=3 รหัสของตู้ 'F' คือเลขอะไร",
            question_en: "Filing cabinets are coded alphabetically: A=1, B=2, C=3. What is the code for cabinet 'F'?",
            answer: "6",
            explanation: "F เป็นตัวอักษรตัวที่ 6 ในภาษาอังกฤษ",
            explanation_en: "F is the 6th letter of the English alphabet."
        },
        {
            id: "F5_C09", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "ในห้องประชุมมีเก้าอี้ 4 แถว แถวละ 8 ตัว มีเก้าอี้ทั้งหมดกี่ตัว",
            question_en: "A conference room has 4 rows of chairs with 8 chairs in each row. How many chairs are there in total?",
            answer: "32",
            explanation: "4 คูณ 8 เท่ากับ 32",
            explanation_en: "4 multiplied by 8 equals 32."
        },
        {
            id: "F5_C10", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "บนปฏิทินของสำนักงาน ถ้าวันที่ 5 ตรงกับวันจันทร์ วันจันทร์ในสัปดาห์ถัดไปคือวันที่เท่าไหร่",
            question_en: "On the office calendar, if the 5th is a Monday, what date is the Monday of the following week?",
            answer: "12",
            explanation: "สัปดาห์หนึ่งมี 7 วัน 5 + 7 = 12",
            explanation_en: "There are 7 days in a week: 5 + 7 = 12."
        },
        {
            id: "F5_C11", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "รหัสตู้เซฟคือเลข 3 หลัก เลขหลักแรกคือ 4 หลักที่สองน้อยกว่าหลักแรก 1 หลักที่สามคือผลรวมของสองหลักแรก รหัสคืออะไร",
            question_en: "A 3-digit safe code has 4 as the first digit. The second digit is 1 less than the first, and the third digit is the sum of the first two. What is the code?",
            answer: "437",
            explanation: "หลักแรก 4, หลักที่สอง 4-1=3, หลักที่สาม 4+3=7 รวมเป็น 437",
            explanation_en: "First digit is 4, second is 4-1=3, third is 4+3=7, forming 437."
        },
        {
            id: "F5_C12", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "ตู้กดน้ำดื่มเหลือน้ำ 20 ลิตร พนักงานกดกินคนละ 2 ลิตร จะกดได้กี่คนก่อนน้ำหมด",
            question_en: "A water dispenser has 20 liters left. If each employee takes 2 liters, how many employees can get water before it runs out?",
            answer: "10",
            explanation: "20 หาร 2 เท่ากับ 10",
            explanation_en: "20 divided by 2 equals 10."
        },
        {
            id: "F5_C13", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "บัตรพนักงานหมายเลข 0012, 0024, 0036 หมายเลขต่อไปคืออะไร",
            question_en: "Employee ID badges are numbered 0012, 0024, 0036. What is the next badge number?",
            answer: "0048",
            explanation: "ตัวเลขเพิ่มขึ้นทีละ 12 ดังนั้น 36 + 12 = 48 (เติม 00 ข้างหน้าตามรูปแบบ)",
            explanation_en: "The numbers increase by 12, so 36 + 12 = 48 (padded to 0048)."
        },
        {
            id: "F5_C14", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "มีปากกา 50 ด้าม เป็นสีแดง 15 ด้าม สีน้ำเงิน 20 ด้าม ที่เหลือเป็นสีดำ มีปากกาสีดำกี่ด้าม",
            question_en: "There are 50 pens: 15 are red, 20 are blue, and the rest are black. How many black pens are there?",
            answer: "15",
            explanation: "50 - (15 + 20) = 15",
            explanation_en: "50 - (15 + 20) = 15."
        },
        {
            id: "F5_C15", floor: 5, theme: "สำนักงานร้าง", difficulty: "ง่ายถึงปานกลาง",
            question: "ลิฟต์รับน้ำหนักได้ 500 กก. หากพนักงานแต่ละคนหนักเฉลี่ย 80 กก. ลิฟต์จะรับพนักงานได้มากที่สุดกี่คนโดยไม่เกินพิกัด",
            question_en: "An elevator has a 500 kg weight limit. If each employee weighs an average of 80 kg, what is the maximum number of employees it can carry safely?",
            answer: "6",
            explanation: "500 หาร 80 ได้ 6.25 ปัดลงเหลือ 6 คน",
            explanation_en: "500 divided by 80 is 6.25, rounded down to 6 people."
        },

        // ==========================================
        // ชั้น 4: ห้องปฏิบัติการ (ปานกลาง)
        // ==========================================
        {
            id: "F4_C01", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "การผสมสาร A และ B ต้องใช้อัตราส่วน 2:5 หากใช้สาร A ปริมาณ 10 ml ต้องใช้สาร B กี่ ml",
            question_en: "Mixing chemicals A and B requires a 2:5 ratio. If you use 10 ml of chemical A, how many ml of chemical B are needed?",
            answer: "25",
            explanation: "อัตราส่วน 2:5, ถ้า A=10 (คูณ 5), ดังนั้น B = 5 * 5 = 25",
            explanation_en: "Ratio 2:5; if A=10 (x5), then B = 5 * 5 = 25."
        },
        {
            id: "F4_C02", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "อุณหภูมิในตู้แช่สารเคมีลดลงครึ่งหนึ่งทุกๆ 2 ชั่วโมง หากตอนเริ่มอุณหภูมิคือ 64 องศา ผ่านไป 6 ชั่วโมง อุณหภูมิจะเป็นเท่าไหร่",
            question_en: "The temperature in a chemical freezer halves every 2 hours. If it starts at 64 degrees, what will the temperature be after 6 hours?",
            answer: "8",
            explanation: "ลดลง 3 ครั้ง (6/2): 64 -> 32 -> 16 -> 8",
            explanation_en: "Halved 3 times (6/2): 64 -> 32 -> 16 -> 8."
        },
        {
            id: "F4_C03", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "รหัสเครื่องปั่นเหวี่ยงคือเลขที่หายไปจากชุดข้อมูล: 3, 6, 12, 24, __, 96",
            question_en: "The centrifuge code is the missing number in the dataset: 3, 6, 12, 24, __, 96.",
            answer: "48",
            explanation: "ตัวเลขแต่ละตัวคูณด้วย 2, ดังนั้น 24 * 2 = 48",
            explanation_en: "Each number is multiplied by 2, so 24 * 2 = 48."
        },
        {
            id: "F4_C04", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "นักวิจัยบันทึกน้ำหนักสารทดลอง: วันที่หนึ่ง 15g, วันที่สอง 12g, วันที่สาม 9g ในวันที่ห้าสารจะหนักกี่กรัม หากน้ำหนักลดลงในอัตราคงที่",
            question_en: "A researcher logs sample weights: Day 1: 15g, Day 2: 12g, Day 3: 9g. At this constant rate of decrease, how many grams will it weigh on Day 5?",
            answer: "3",
            explanation: "ลดลงวันละ 3g, วันที่สี่ = 6g, วันที่ห้า = 3g",
            explanation_en: "Decreases by 3g per day: Day 4 = 6g, Day 5 = 3g."
        },
        {
            id: "F4_C05", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "ขวดสารเคมีมีตัวเลขกำกับ 11, 13, 17, 19, 23, __ ขวดถัดไปคือหมายเลขอะไร",
            question_en: "Chemical vials are labeled 11, 13, 17, 19, 23, __. What is the number on the next vial?",
            answer: "29",
            explanation: "เป็นลำดับของจำนวนเฉพาะ (Prime numbers)",
            explanation_en: "It is a sequence of prime numbers."
        },
        {
            id: "F4_C06", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "มีหลอดทดลอง 120 หลอด แต่ละชั้นวางได้ 15 หลอด ต้องใช้ชั้นวางกี่ชั้นจึงจะวางหลอดทดลองได้ทั้งหมดพอดี",
            question_en: "There are 120 test tubes, and each rack holds 15 tubes. How many racks are needed to hold all the test tubes?",
            answer: "8",
            explanation: "120 หาร 15 เท่ากับ 8",
            explanation_en: "120 divided by 15 equals 8."
        },
        {
            id: "F4_C07", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "สาร X 1 หยด มีปริมาตร 0.05 ml หากต้องการสาร 2 ml ต้องใช้กี่หยด",
            question_en: "One drop of Substance X has a volume of 0.05 ml. How many drops are needed to get 2 ml?",
            answer: "40",
            explanation: "2 หาร 0.05 เท่ากับ 40",
            explanation_en: "2 divided by 0.05 equals 40."
        },
        {
            id: "F4_C08", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "รหัสเปิดตู้ปลอดเชื้อคือเลขหลักหน่วยของผลลัพธ์ 7 ยกกำลัง 3 รหัสคืออะไร",
            question_en: "The sterile cabinet unlock code is the units (ones) digit of 7 to the power of 3. What is the code?",
            answer: "3",
            explanation: "7 * 7 = 49, 49 * 7 = 343 เลขหลักหน่วยคือ 3",
            explanation_en: "7 * 7 = 49, 49 * 7 = 343; the units digit is 3."
        },
        {
            id: "F4_C09", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "เครื่องชั่งตัวแรก อ่านค่าได้ 150g เครื่องชั่งที่2 อ่านได้มากกว่าตัวแรก 10% เครื่องชั่งที่2  อ่านค่าได้เท่าไหร่",
            question_en: "The first scale reads 150g. The second scale reads 10% more than the first. What does the second scale read?",
            answer: "165",
            explanation: "10% ของ 150 คือ 15 นำไปบวก 150 ได้ 165",
            explanation_en: "10% of 150 is 15; adding it to 150 gives 165."
        },
        {
            id: "F4_C10", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "บันทึกห้องแล็บระบุ: 'รหัสคือผลคูณของจำนวนขาของแมงมุมและจำนวนขาของแมลงทั่วไป' รหัสคืออะไร",
            question_en: "A lab note states: 'The code is the product of the number of legs on a spider and the number of legs on a typical insect.' What is the code?",
            answer: "48",
            explanation: "แมงมุมมี 8 ขา แมลงมี 6 ขา 8 * 6 = 48",
            explanation_en: "A spider has 8 legs and an insect has 6 legs: 8 * 6 = 48."
        },
        {
            id: "F4_C11", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "ปฏิกิริยาเคมีใช้เวลา 1 ชั่วโมง 15 นาที หากเริ่มตอน 14:50 น. จะเสร็จสิ้นเวลาใด (ตอบ HH:MM)",
            question_en: "A chemical reaction takes 1 hour and 15 minutes. If it starts at 14:50, what time will it finish? (Answer in HHMM format)",
            answer: "1605",
            explanation: "14:50 บวก 1 ชม. = 15:50 บวกอีก 15 นาที = 16:05",
            explanation_en: "14:50 plus 1 hour = 15:50, plus 15 minutes = 16:05."
        },
        {
            id: "F4_C12", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "กล้องจุลทรรศน์ใช้เลนส์ใกล้วัตถุ 40X และเลนส์ใกล้ตา 10X กำลังขยายรวมเป็นกี่เท่า",
            question_en: "A microscope uses a 40X objective lens and a 10X eyepiece lens. What is the total magnification?",
            answer: "400",
            explanation: "กำลังขยายรวม = 40 * 10 = 400",
            explanation_en: "Total magnification = 40 * 10 = 400."
        },
        {
            id: "F4_C13", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "มีหนูทดลอง 5 กรง กรงละ 4 ตัว หากย้ายหนูทั้งหมดมารวมในกรงใหม่ 2 กรง ให้เท่ากัน จะมีกรงละกี่ตัว",
            question_en: "There are 5 cages with 4 lab mice each. If all mice are moved equally into 2 new cages, how many mice will be in each new cage?",
            answer: "10",
            explanation: "หนูทั้งหมด 5 * 4 = 20 ตัว แบ่ง 2 กรง กรงละ 10 ตัว",
            explanation_en: "Total mice = 5 * 4 = 20; divided into 2 cages gives 10 per cage."
        },
        {
            id: "F4_C14", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "รหัสปลดล็อกสารเคมีอันตรายคือค่าเฉลี่ยของตัวเลข 12, 18 และ 30 รหัสคืออะไร",
            question_en: "The hazardous chemical unlock code is the average of 12, 18, and 30. What is the code?",
            answer: "20",
            explanation: "(12 + 18 + 30) / 3 = 60 / 3 = 20",
            explanation_en: "(12 + 18 + 30) / 3 = 60 / 3 = 20."
        },
        {
            id: "F4_C15", floor: 4, theme: "ห้องปฏิบัติการ", difficulty: "ปานกลาง",
            question: "ความยาวคลื่นของแสงสีหนึ่งถูกเข้ารหัสเป็น: นำเลข 500 บวก 200 แล้วหักออก 10% รหัสนั้นคือเท่าไหร่",
            question_en: "A light wavelength is encoded as: add 500 and 200, then subtract 10%. What is the resulting code?",
            answer: "630",
            explanation: "500 + 200 = 700 หัก 10% (คือ 70) เหลือ 630",
            explanation_en: "500 + 200 = 700; subtracting 10% (70) leaves 630."
        },

        // ==========================================
        // ชั้น 3: เขตเครื่องจักร (ปานกลางถึงยาก)
        // ==========================================
        {
            id: "F3_C01", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "เฟือง A มี 20 ซี่ ขบกับเฟือง B ที่มี 10 ซี่ หากเฟือง A หมุน 5 รอบ เฟือง B จะหมุนกี่รอบ",
            question_en: "Gear A has 20 teeth and meshes with Gear B, which has 10 teeth. If Gear A rotates 5 times, how many times will Gear B rotate?",
            answer: "10",
            explanation: "เฟือง B เล็กกว่าครึ่งหนึ่ง จึงต้องหมุนเป็น 2 เท่า ดังนั้น 5 * 2 = 10",
            explanation_en: "Gear B is half the size, so it rotates twice as fast: 5 * 2 = 10."
        },
        {
            id: "F3_C02", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "ท่อน้ำมีแรงดัน 500 PSI วาล์วตัวหนึ่งช่วยลดแรงดันได้ 20% หากต่อวาล์วนี้ 2 ตัวเรียงกัน แรงดันสุดท้ายจะเป็นเท่าไหร่",
            question_en: "A water pipe has 500 PSI of pressure. One valve reduces pressure by 20%. If two such valves are connected in series, what is the final pressure?",
            answer: "320",
            explanation: "ลด 20% ครั้งแรก: 500 * 0.8 = 400, ลด 20% ครั้งที่สอง: 400 * 0.8 = 320",
            explanation_en: "First 20% reduction: 500 * 0.8 = 400; second 20% reduction: 400 * 0.8 = 320."
        },
        {
            id: "F3_C03", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "สายไฟ 4 เส้นมีหมายเลขกำกับ 10, 20, 30, 40 หากต้องตัดสายไฟที่มีผลรวมของตัวเลขเท่ากับ 50 โดยใช้สายไฟเพียง 2 เส้น เส้นที่มีหมายเลขน้อยกว่าคือเลขอะไร",
            question_en: "Four wires are numbered 10, 20, 30, and 40. You must cut 2 wires whose numbers sum to 50. What is the number of the smaller wire?",
            answer: "10",
            explanation: "10 + 40 = 50 หมายเลขที่น้อยกว่าคือ 10",
            explanation_en: "10 + 40 = 50; the smaller wire number is 10."
        },
        {
            id: "F3_C04", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "รหัสเปิดแผงควบคุมหลักคือตัวเลขที่ทำให้สมการ 3x - 15 = 45 เป็นจริง ค่าของ x คืออะไร",
            question_en: "The main control panel code is the value of x that satisfies 3x - 15 = 45. What is x?",
            answer: "20",
            explanation: "3x = 60, x = 20",
            explanation_en: "3x = 60, so x = 20."
        },
        {
            id: "F3_C05", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "เครื่องกำเนิดไฟฟ้าต้องการน้ำมัน 3 ลิตรต่อชั่วโมง หากมีน้ำมันอยู่ 45 ลิตร แต่เครื่องรั่วทำให้น้ำมันหายไป 1.5 ลิตรต่อชั่วโมง เครื่องจะทำงานได้นานกี่ชั่วโมง",
            question_en: "A generator consumes 3 liters of fuel per hour. With 45 liters available and a leak losing 1.5 liters per hour, how many hours can it run?",
            answer: "10",
            explanation: "ใช้น้ำมันรวม 3 + 1.5 = 4.5 ลิตร/ชม. ดังนั้น 45 / 4.5 = 10 ชั่วโมง",
            explanation_en: "Total fuel loss is 3 + 1.5 = 4.5 liters/hr, so 45 / 4.5 = 10 hours."
        },
        {
            id: "F3_C06", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "รหัส 3 หลัก: ผลรวมของทั้งสามหลักคือ 12 หลักแรกเป็นสองเท่าของหลักสุดท้าย หลักที่สองคือ 3 รหัสนั้นคืออะไร",
            question_en: "3-digit code: The sum of all three digits is 12. The first digit is double the last digit, and the second digit is 3. What is the code?",
            answer: "633",
            explanation: "หลักที่สอง=3 ผลรวมที่เหลือ=9 หลักแรกต้องเป็น 2 เท่าของหลักสุดท้าย ดังนั้นหลักแรก=6 หลักสุดท้าย=3",
            explanation_en: "Second digit = 3, remaining sum = 9. First digit is double the last, so first = 6 and last = 3 (633)."
        },
        {
            id: "F3_C07", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "สวิตช์ไฟเรียงเป็นแถว 100 ตัว สวิตช์ถูกสับเปลี่ยนสถานะทุกๆ ตัวที่หารด้วย 3 ลงตัว จะมีสวิตช์กี่ตัวที่ถูกสับเปลี่ยนสถานะ",
            question_en: "There are 100 light switches in a row. Every switch divisible by 3 is toggled. How many switches are toggled?",
            answer: "33",
            explanation: "100 หาร 3 ได้ 33 เศษ 1 ดังนั้นมี 33 ตัว",
            explanation_en: "100 divided by 3 is 33 with a remainder of 1, so 33 switches."
        },
        {
            id: "F3_C08", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "ลิฟต์ขนของเคลื่อนที่ด้วยความเร็ว 2 เมตรต่อวินาที ระยะทางจากชั้น 1 ถึงชั้น 3 คือ 30 เมตร ลิฟต์ใช้เวลากี่วินาที",
            question_en: "A freight elevator moves at 2 meters per second. The distance from Floor 1 to Floor 3 is 30 meters. How many seconds does it take?",
            answer: "15",
            explanation: "30 หาร 2 เท่ากับ 15 วินาที",
            explanation_en: "30 divided by 2 equals 15 seconds."
        },
        {
            id: "F3_C09", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "แบตเตอรี่ชาร์จเต็ม 100% ใช้งานไป 1/4 ของความจุ จากนั้นชาร์จเพิ่มอีก 10% ของที่เหลืออยู่ แบตเตอรี่จะมีพลังงานกี่เปอร์เซ็นต์ ตอบโดยไม่ต้องใส่ทศนิยม",
            question_en: "A battery is at 100%, uses 1/4 of its capacity, and is then charged by 10% of its remaining level. What percentage is it now? (Enter without decimal point)",
            answer: "825",
            explanation: "เหลือ 75%, ชาร์จเพิ่ม 10% ของ 75 คือ 7.5%, รวมเป็น 75 + 7.5 = 82.5",
            explanation_en: "75% remains; adding 10% of 75 (7.5%) gives 82.5%, entered without decimal as 825."
        },
        {
            id: "F3_C10", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "รหัสประตูห้องเครื่องเป็นลำดับตัวเลข: 13, 15, 17, 19 ตัวเลขถัดไปคืออะไร",
            question_en: "The engine room door code follows the sequence: 13, 15, 17, 19. What is the next number?",
            answer: "21",
            explanation: "ลำดับตัวเลขเพิ่มขึ้นทีละ 2 ดังนั้น 19 + 2 = 21",
            explanation_en: "The sequence increases by 2, so 19 + 2 = 21."
        },
        {
            id: "F3_C11", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "ปั๊มน้ำ A เติมถังเต็มใน 3 ชั่วโมง ปั๊ม B เติมเต็มใน 6 ชั่วโมง หากเปิดพร้อมกันจะใช้เวลากี่ชั่วโมง",
            question_en: "Water Pump A fills a tank in 3 hours, while Pump B fills it in 6 hours. If both run together, how many hours will it take?",
            answer: "2",
            explanation: "ความเร็วรวม 1/3 + 1/6 = 3/6 = 1/2 ถังต่อชั่วโมง ดังนั้นใช้เวลา 2 ชั่วโมง",
            explanation_en: "Combined rate: 1/3 + 1/6 = 1/2 tank per hour, so it takes 2 hours."
        },
        {
            id: "F3_C12", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "มอเตอร์หมุน 1500 รอบต่อนาที หากผ่านไป 12 วินาที มอเตอร์หมุนไปแล้วกี่รอบ",
            question_en: "A motor spins at 1,500 revolutions per minute. How many revolutions does it complete in 12 seconds?",
            answer: "300",
            explanation: "1500 รอบ/60 วินาที = 25 รอบ/วินาที, 25 * 12 = 300 รอบ",
            explanation_en: "1,500 rev / 60 sec = 25 rev/sec; 25 * 12 = 300 revolutions."
        },
        {
            id: "F3_C13", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "หาก A = 1, B = 3, C = 6, D = 10 แล้ว E จะมีค่าเท่าใดตามรูปแบบนี้",
            question_en: "If A = 1, B = 3, C = 6, and D = 10, what is the value of E following this pattern?",
            answer: "15",
            explanation: "เพิ่มขึ้นทีละ +2, +3, +4, ดังนั้น E คือ 10 + 5 = 15",
            explanation_en: "Increases by +2, +3, +4, so E is 10 + 5 = 15."
        },
        {
            id: "F3_C14", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "รอกหนึ่งตัวผ่อนแรงได้ครึ่งหนึ่ง หากต้องยกของหนัก 200 กก. โดยใช้รอก 3 ตัวซ้อนกัน ต้องออกแรงเทียบเท่ายกของกี่กิโลกรัม",
            question_en: "A single pulley halves the required lifting force. If you lift a 200 kg load using 3 pulleys in series, how many kilograms of force are needed?",
            answer: "25",
            explanation: "200 -> 100 -> 50 -> 25",
            explanation_en: "200 -> 100 -> 50 -> 25 kg."
        },
        {
            id: "F3_C15", floor: 3, theme: "เขตเครื่องจักร", difficulty: "ปานกลางถึงยาก",
            question: "หน้าปัดเกจวัดชี้ที่เลข 8 จากทั้งหมด 12 หากเข็มหมุนตามเข็มนาฬิกาไป 240 องศา เข็มจะชี้ที่เลขใด",
            question_en: "A gauge dial points to 8 out of 12. If the needle rotates clockwise by 240 degrees, what number will it point to?",
            answer: "4",
            explanation: "หน้าปัด 360 องศาแบ่งเป็น 12 ช่อง ช่องละ 30 องศา, 240/30 = 8 ช่อง, 8 + 8 = 16 หัก 12 เหลือ 4",
            explanation_en: "360 degrees / 12 = 30 degrees per step; 240/30 = 8 steps clockwise from 8 reaches 4."
        },

        // ==========================================
        // ชั้น 2: ศูนย์รักษาความปลอดภัย (ยาก)
        // ==========================================
        {
            id: "F2_C01", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "ยาม 4 คนมีรหัสประจำตัวคือ 101, 102, 103, 104 หากนำรหัสประจำตัวของคนที่รหัสมากที่สุดลบด้วยรหัสที่น้อยที่สุด จะได้ผลลัพธ์เท่าใด",
            question_en: "Four guards have ID codes 101, 102, 103, and 104. Subtract the smallest ID code from the largest ID code. What is the result?",
            answer: "3",
            explanation: "104 - 101 = 3",
            explanation_en: "104 - 101 = 3."
        },
        {
            id: "F2_C02", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "กล้องวงจรปิดหมุนไปทางซ้าย 90 องศา จากนั้นหมุนไปทางขวา 180 องศา และหมุนไปทางซ้ายอีก 270 องศา กล้องหมุนไปจากจุดเริ่มต้นรวมกี่องศา (คิดเฉพาะขนาดมุมที่เปลี่ยนแปลงไปจากจุดเริ่มต้น)",
            question_en: "A security camera turns left 90 degrees, then right 180 degrees, and then left 270 degrees. By how many degrees has its angle changed from the start?",
            answer: "180",
            explanation: "ซ้าย 90(-90), ขวา 180(+180), ซ้าย 270(-270) รวมเป็น -180 องศา หรือต่างจากจุดเริ่มต้น 180 องศา",
            explanation_en: "Left 90 (-90), right 180 (+180), left 270 (-270) = -180 degrees, a 180-degree difference."
        },
        {
            id: "F2_C03", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "รหัสปลดล็อกปืน: เลข 4 หลัก ผลคูณของทุกหลักคือ 0 ผลรวมคือ 9 หลักแรกคือ 5 หลักที่สองและสามเหมือนกัน รหัสคืออะไร",
            question_en: "Weapon lock code: 4 digits. The product of all digits is 0, the sum is 9, the first digit is 5, and the second and third digits are identical. What is the code?",
            answer: "5220",
            explanation: "ต้องมีเลข 0 หนึ่งตัว หลักแรกคือ 5 เหลือผลรวม 4 สำหรับสองหลักที่เหมือนกัน (2+2) ดังนั้นคือ 5220",
            explanation_en: "Must contain 0; first digit is 5, leaving a sum of 4 for the two identical middle digits (2+2), giving 5220."
        },
        {
            id: "F2_C04", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "เซิร์ฟเวอร์สำรองข้อมูลทุก 45 นาที หากเริ่มครั้งแรกตอน 00:00 น. การสำรองข้อมูลครั้งที่ 10 จะเกิดขึ้นเวลาใด (ตอบ HH:MM)",
            question_en: "The server backs up data every 45 minutes. If the first backup is at 00:00, what time will the 10th backup occur? (Answer in HHMM format)",
            answer: "0645",
            explanation: "ครั้งที่ 1 คือ 00:00, ห่างกัน 9 ช่วงเวลา 9 x 45 นาที = 405 นาที (6 ชั่วโมง 45 นาที) ตรงกับ 06:45",
            explanation_en: "1st is at 00:00; 9 intervals later: 9 * 45 = 405 minutes (6 hrs 45 mins), which is 06:45."
        },
        {
            id: "F2_C05", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "ระบบบันทึกความปลอดภัยตรวจจับผู้บุกรุกได้ที่ประตูหมายเลข 1 หรือ 2 หากเซ็นเซอร์ประตูหมายเลข 1 ขัดข้องและแสดงค่าเป็นศูนย์ตลอดเวลา ประตูที่ทำงานและตรวจจับได้คือประตูหมายเลขอะไร",
            question_en: "The security log detects intruders at Door 1 or Door 2. If Door 1's sensor is broken and always reads zero, which door number detected the intruder?",
            answer: "2",
            explanation: "ประตูหมายเลข 1 ขัดข้อง จึงต้องเป็นประตูหมายเลข 2",
            explanation_en: "Door 1 is malfunctioning, so it must be Door 2."
        },
        {
            id: "F2_C06", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "รหัสวิทยุสื่อสารรับสัญญาณเป็นตัวเลข 8, 15, 15 หากต้องถอดรหัสโดยลดค่าตัวเลขแต่ละตัวลง 3 ผลรวมของตัวเลขที่ถอดรหัสแล้วคือเท่าใด",
            question_en: "A radio receives the numbers 8, 15, 15. To decode them, subtract 3 from each number. What is the sum of the decoded numbers?",
            answer: "29",
            explanation: "8-3=5, 15-3=12, 15-3=12 ผลรวมคือ 5+12+12 = 29",
            explanation_en: "8-3=5, 15-3=12, 15-3=12; their sum is 5 + 12 + 12 = 29."
        },
        {
            id: "F2_C07", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "มีเซ็นเซอร์ 5 ตัว (1,2,3,4,5) ติดตั้งเรียงกัน เซ็นเซอร์คู่ไม่ทำงาน ตัวที่ 1 และ 5 เป็นปกติ เซ็นเซอร์ตัวใดบ้างที่ทำงาน (ตอบเป็นตัวเลขติดกันไม่มีช่องว่าง)",
            question_en: "Five sensors (1, 2, 3, 4, 5) are installed in a row. Even-numbered sensors are offline, while 1 and 5 are normal. Which sensors are active? (Enter digits without spaces)",
            answer: "135",
            explanation: "เลขคู่ 2, 4 ไม่ทำงาน ดังนั้นเหลือ 1, 3, 5",
            explanation_en: "Even sensors 2 and 4 are offline, leaving 1, 3, and 5 (135)."
        },
        {
            id: "F2_C08", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "รหัสผ่านห้องคุมขังคือผลรวมของจำนวนเฉพาะทั้งหมดที่อยู่ระหว่าง 10 ถึง 20 รหัสนั้นคืออะไร",
            question_en: "The holding cell passcode is the sum of all prime numbers between 10 and 20. What is the passcode?",
            answer: "60",
            explanation: "จำนวนเฉพาะคือ 11, 13, 17, 19 รวมกันได้ 60",
            explanation_en: "The primes are 11, 13, 17, and 19, which sum to 60."
        },
        {
            id: "F2_C09", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "เจ้าหน้าที่ 4 คนมีระดับสิทธิ์การเข้าถึงดังนี้: คนที่ 1 ระดับ 5, คนที่ 2 ระดับ 3, คนที่ 3 ระดับ 2, คนที่ 4 ระดับ 6 ผลรวมของระดับสิทธิ์ที่ต่ำที่สุดและสูงที่สุดคือเท่าใด",
            question_en: "Four officers have access levels: Officer 1: Level 5, Officer 2: Level 3, Officer 3: Level 2, Officer 4: Level 6. What is the sum of the lowest and highest levels?",
            answer: "8",
            explanation: "ระดับต่ำสุดคือ 2 สูงสุดคือ 6 รวมกันได้ 2 + 6 = 8",
            explanation_en: "Lowest is 2 and highest is 6; 2 + 6 = 8."
        },
        {
            id: "F2_C10", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "ระบบตรวจสอบลายนิ้วมือให้คะแนนความแม่นยำ 3 ครั้งคือ 95, 88 และ 90 ค่าเฉลี่ยของคะแนนทั้งสามครั้งคือเท่าใด",
            question_en: "A fingerprint scanner gave three accuracy scores: 95, 88, and 90. What is the average of these three scores?",
            answer: "91",
            explanation: "(95 + 88 + 90) / 3 = 273 / 3 = 91",
            explanation_en: "(95 + 88 + 90) / 3 = 273 / 3 = 91."
        },
        {
            id: "F2_C11", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "นาฬิกาในศูนย์ รปภ. เดินเร็ว 5 นาทีต่อชั่วโมง หากตั้งเวลาตรงกันตอน 12:00 น. เมื่อเวลาจริงคือ 18:00 น. นาฬิกาเรือนนี้จะบอกเวลาเท่าใด (ตอบ HH:MM)",
            question_en: "A security room clock runs 5 minutes fast every hour. Set accurately at 12:00, what time will this clock show when the real time is 18:00? (Answer in HHMM format)",
            answer: "1830",
            explanation: "ผ่านไป 6 ชั่วโมง เดินเร็วขึ้น 6 * 5 = 30 นาที เป็น 18:30",
            explanation_en: "After 6 hours, it gains 6 * 5 = 30 minutes, showing 18:30."
        },
        {
            id: "F2_C12", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "รหัสคีย์การ์ด: เลขหลักแรกคือ 8 เลขหลักที่สองคือครึ่งหนึ่งของหลักแรก เลขหลักที่สามคือผลคูณของสองหลักแรก และหลักที่สี่คือหลักที่สามลบ 30 รหัสคืออะไร",
            question_en: "Keycard code: The 1st part is 8, the 2nd is half of the 1st, the 3rd is the product of the first two, and the 4th is the 3rd minus 30. What is the code?",
            answer: "84322",
            explanation: "หลักที่ 1=8, หลักที่ 2=4, ส่วนที่สาม=32, ส่วนที่สี่=2 รวมกันเป็นสตริง 84322",
            explanation_en: "1st = 8, 2nd = 4, 3rd = 32, 4th = 2; concatenated together as 84322."
        },
        {
            id: "F2_C13", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "รหัสประตูฉุกเฉินคือตัวเลข x ที่ทำให้ (x/2) + 10 = 25 รหัสคืออะไร",
            question_en: "The emergency door code is the number x where (x/2) + 10 = 25. What is the code?",
            answer: "30",
            explanation: "(x/2) = 15, x = 30",
            explanation_en: "(x/2) = 15, so x = 30."
        },
        {
            id: "F2_C14", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "มีประตู 3 บาน: บาน 1 เป็นทางตัน, บาน 2 เป็นกับดัก, บาน 3 ปลอดภัย ป้ายหน้าบาน 2 เขียนว่า 'บาน 1 ปลอดภัย' ถ้าป้ายนี้โกหก ประตูใดปลอดภัย (ตอบเป็นหมายเลข 1, 2 หรือ 3)",
            question_en: "There are 3 doors: Door 1 is a dead end, Door 2 is a trap, and Door 3 is safe. The sign on Door 2 says 'Door 1 is safe.' If the sign is lying, which door is safe? (1, 2, or 3)",
            answer: "3",
            explanation: "จากข้อมูล บาน 3 ปลอดภัยอยู่แล้ว ป้ายโกหกเพียงยืนยันว่าบาน 1 ไม่ปลอดภัย ดังนั้นคำตอบคือบาน 3",
            explanation_en: "Door 3 is safe; the lying sign simply confirms Door 1 is not safe, so the answer is 3."
        },
        {
            id: "F2_C15", floor: 2, theme: "ศูนย์รักษาความปลอดภัย", difficulty: "ยาก",
            question: "ไฟล์บันทึกภาพถูกแบ่งเป็น 8 ส่วนเท่าๆ กัน หากแต่ละส่วนมีขนาด 1.5 GB และเซิร์ฟเวอร์ดาวน์โหลดได้ 2 GB ต่อวินาที ต้องใช้เวลากี่วินาทีในการดาวน์โหลดทั้งหมด",
            question_en: "A video log is split into 8 equal parts of 1.5 GB each. If the server downloads at 2 GB per second, how many seconds will it take to download everything?",
            answer: "6",
            explanation: "ขนาดรวม 8 * 1.5 = 12 GB, 12 / 2 = 6 วินาที",
            explanation_en: "Total size = 8 * 1.5 = 12 GB; 12 / 2 = 6 seconds."
        },

        // ==========================================
        // ชั้น 1: ประตูทางออกสุดท้าย (ยากที่สุด)
        // ==========================================
        {
            id: "F1_C01", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "แผงควบคุมหลักมี 4 ปุ่มหมายเลข 1, 2, 3, 4 กฎคือ: ผลบวกของปุ่มที่กดสองอันดับแรกต้องเท่ากับ 3 และปุ่มสุดท้ายที่ต้องกดคือหมายเลข 4 ปุ่มแรกที่ต้องกดคือหมายเลขใด หากหมายเลข 2 ถูกกดเป็นอันดับที่สอง",
            question_en: "The main panel has 4 buttons: 1, 2, 3, 4. The sum of the first two buttons pressed must equal 3, and button 4 is pressed last. If button 2 is pressed second, which button must be pressed first?",
            answer: "1",
            explanation: "ผลบวกสองอันดับแรกคือ 3 กดปุ่ม 2 ไปแล้ว ดังนั้นปุ่มแรกคือ 1",
            explanation_en: "The sum of the first two is 3; since 2 is pressed second, the first button must be 1."
        },
        {
            id: "F1_C02", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "รหัส 5 หลัก: หลักกลางคือ 5 หลักแรกน้อยกว่าหลักกลาง 2 หลักสุดท้ายมากกว่าหลักกลาง 3 หลักที่สองคือผลรวมของหลักแรกและหลักกลาง หลักที่สี่คือ 0 รหัสคืออะไร",
            question_en: "5-digit code: The middle digit is 5. The 1st is 2 less than the middle, the last is 3 more than the middle, the 2nd is the sum of the 1st and middle, and the 4th is 0. What is the code?",
            answer: "38508",
            explanation: "หลักกลาง(3)=5, หลักแรก(1)=3, หลักสุดท้าย(5)=8, หลักที่สอง(2)=3+5=8, หลักสี่(4)=0 -> 38508",
            explanation_en: "Middle(3rd)=5, 1st=3, 5th=8, 2nd=3+5=8, 4th=0 -> 38508."
        },
        {
            id: "F1_C03", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "คุณมีขวดน้ำ 3 ลิตร และ 5 ลิตร หากต้องการตวงน้ำให้ได้ 4 ลิตรพอดี ต้องเทน้ำ 3 ลิตรใส่ขวด 5 ลิตรกี่ครั้ง (เทจนกว่าจะได้ 4 ลิตรในขวด 5 ลิตร)",
            question_en: "You have a 3-liter bottle and a 5-liter bottle. To measure exactly 4 liters into the 5-liter bottle, how many times do you fill the 3-liter bottle from the tap?",
            answer: "2",
            explanation: "เท 3 ลิตรลง 5 ลิตร (ครั้งที่ 1), เท 3 ลิตรลง 5 ลิตรอีกจนเต็ม (ขวด 5 ลิตรเต็ม, เหลือในขวด 3 ลิตร 1 ลิตร) ทิ้งน้ำ 5 ลิตร เท 1 ลิตรใส่ขวด 5 ลิตร แล้วเท 3 ลิตรลงขวด 5 ลิตร (ครั้งที่ 2) รวมตวงจากขวด 3 ลิตร 2 ครั้ง (แก้ไข: ต้องตวงจากก๊อกใส่ 3 ลิตร 2 ครั้ง)",
            explanation_en: "Fill 3L and pour into 5L (1st fill), fill 3L again and top off 5L leaving 1L in 3L (2nd fill), empty 5L, pour 1L in, then pour 3L in = 4L (2 fills of the 3L bottle)."
        },
        {
            id: "F1_C04", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "หาก A+B=10, B+C=15, A+C=17 แล้ว A+B+C เท่ากับเท่าใด",
            question_en: "If A + B = 10, B + C = 15, and A + C = 17, what does A + B + C equal?",
            answer: "21",
            explanation: "รวมสมการทั้งหมด 2A+2B+2C = 42 ดังนั้น A+B+C = 21",
            explanation_en: "Adding all equations gives 2A + 2B + 2C = 42, so A + B + C = 21."
        },
        {
            id: "F1_C05", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "หากนำตัวเลขวันในสัปดาห์ (อาทิตย์=1, จันทร์=2, ... เสาร์=7) ของวันศุกร์มาคูณด้วยวันอาทิตย์ จะได้ผลลัพธ์เท่าใด",
            question_en: "If days of the week are numbered Sunday=1, Monday=2, ..., Saturday=7, what do you get when you multiply Friday by Sunday?",
            answer: "6",
            explanation: "วันศุกร์ = 6, วันอาทิตย์ = 1 ดังนั้น 6 * 1 = 6",
            explanation_en: "Friday = 6 and Sunday = 1, so 6 * 1 = 6."
        },
        {
            id: "F1_C06", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ลำดับ Fibonacci: 1, 1, 2, 3, 5, 8, ... เลขตัวที่ 10 ในลำดับนี้คืออะไร",
            question_en: "Fibonacci sequence: 1, 1, 2, 3, 5, 8, ... What is the 10th number in this sequence?",
            answer: "55",
            explanation: "1,1,2,3,5,8,13,21,34,55",
            explanation_en: "1, 1, 2, 3, 5, 8, 13, 21, 34, 55."
        },
        {
            id: "F1_C07", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ห้อง 1 ปลอดภัย ห้อง 2 มีก๊าซพิษ ห้อง 3 มีระเบิด ป้ายหน้าห้อง 1 เขียนว่า 'ห้องนี้มีระเบิด' ป้ายหน้าห้อง 2 เขียนว่า 'ห้อง 1 ปลอดภัย' ป้ายห้อง 3 เขียนว่า 'ก๊าซพิษอยู่ในห้องนี้' หากมีป้ายเดียวที่พูดความจริง ป้ายหน้าห้องใดพูดความจริง (ตอบ 1, 2 หรือ 3)",
            question_en: "Room 1 is safe, Room 2 has poison gas, Room 3 has a bomb. Sign 1 says 'This room has a bomb', Sign 2 says 'Room 1 is safe', Sign 3 says 'Poison gas is in here.' If only one sign tells the truth, which room's sign is true? (1, 2, or 3)",
            answer: "2",
            explanation: "ห้อง 1 ปลอดภัยจริง ดังนั้นป้าย 2 พูดความจริง ป้าย 1 (ระเบิด) โกหก ป้าย 3 (ก๊าซพิษ) โกหกเพราะก๊าซอยู่ห้อง 2",
            explanation_en: "Room 1 is safe, so Sign 2 is true, while Signs 1 and 3 are false."
        },
        {
            id: "F1_C08", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "รหัส 4 หลักเรียงมากไปน้อย ไม่มีเลขซ้ำ ผลรวม 18 เลขสุดท้ายคือ 1 หลักแรกคือ 7 หลักที่สามคือ 4 รหัสคืออะไร",
            question_en: "A 4-digit code is in descending order with no repeated digits, summing to 18. The 1st digit is 7, the 3rd is 4, and the last is 1. What is the code?",
            answer: "7641",
            explanation: "หลักแรก 7, หลักสาม 4, หลักสี่ 1 รวมเป็น 12 เหลือหลักสองคือ 6 รวม 18 พอดี",
            explanation_en: "1st=7, 3rd=4, 4th=1 (sum=12); the remaining 2nd digit is 6 to make 18 (7641)."
        },
        {
            id: "F1_C09", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ในการแข่งขันวิ่งหนีตาย คุณวิ่งแซงคนที่อยู่ในอันดับที่สอง ตอนนี้คุณอยู่อันดับที่เท่าไหร่",
            question_en: "In a race for survival, you overtake the person in 2nd place. What place are you in now?",
            answer: "2",
            explanation: "เมื่อคุณแซงที่สอง คุณก็เข้ามาแทนที่ในอันดับที่สอง ไม่ใช่อันดับหนึ่ง",
            explanation_en: "When you overtake the person in 2nd place, you take their 2nd place spot."
        },
        {
            id: "F1_C10", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ระบบนับถอยหลังเริ่มที่ 1000 และลดลงทีละ 7 ทุกๆ 1 วินาที ตัวเลขจำนวนเต็มบวกที่น้อยที่สุดที่ระบบจะแสดงก่อนจะติดลบคือเลขอะไร",
            question_en: "A countdown starts at 1000 and decreases by 7 every second. What is the smallest positive integer displayed before it goes negative?",
            answer: "6",
            explanation: "1000 หาร 7 ได้ 142 เศษ 6 ตัวเลขสุดท้ายที่เป็นบวกคือ 6",
            explanation_en: "1000 divided by 7 is 142 with a remainder of 6, so 6 is the smallest positive number."
        },
        {
            id: "F1_C11", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "รหัสสุดท้ายแปลงจากฐานสอง 10110 เป็นเลขฐานสิบ รหัสคืออะไร",
            question_en: "Convert the binary code 10110 into a decimal number. What is the code?",
            answer: "22",
            explanation: "(1*16) + (0*8) + (1*4) + (1*2) + (0*1) = 16 + 4 + 2 = 22",
            explanation_en: "(1*16) + (0*8) + (1*4) + (1*2) + (0*1) = 16 + 4 + 2 = 22."
        },
        {
            id: "F1_C12", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ประตูนิรภัย 3 ชั้น แต่ละชั้นต้องการรหัสผ่านที่เพิ่มขึ้นเป็นสองเท่าของชั้นก่อนหน้า หากชั้นแรกใช้รหัส 12 รหัสของชั้นที่สามคืออะไร",
            question_en: "A 3-layer security gate requires a passcode that doubles with each layer. If the first layer uses 12, what is the passcode for the third layer?",
            answer: "48",
            explanation: "ชั้นแรก 12, ชั้นสอง 24, ชั้นสาม 48",
            explanation_en: "Layer 1 is 12, Layer 2 is 24, Layer 3 is 48."
        },
        {
            id: "F1_C13", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "มีประตู 100 บาน ปิดอยู่ทั้งหมด เดินรอบที่ 1 เปิดทุกบาน รอบที่ 2 ปิดบานที่เป็นเลขคู่ รอบที่ 3 สลับสถานะบานที่หารด้วย 3 ลงตัว ประตูบานที่ 6 มีสถานะสลับเปลี่ยนไปกี่ครั้งหลังจากจบ 3 รอบ",
            question_en: "100 doors are all closed. Pass 1 opens all doors. Pass 2 closes even-numbered doors. Pass 3 toggles doors divisible by 3. How many times does Door 6 change state after 3 passes?",
            answer: "3",
            explanation: "รอบ 1 สลับจากปิดเป็นเปิด (ครั้งที่ 1), รอบ 2 หาร 2 ลงตัวสลับจากเปิดเป็นปิด (ครั้งที่ 2), รอบ 3 หาร 3 ลงตัวสลับจากปิดเป็นเปิด (ครั้งที่ 3) รวมสลับ 3 ครั้ง",
            explanation_en: "Pass 1 opens it (1), Pass 2 closes it (2), Pass 3 opens it (3) -> 3 state changes."
        },
        {
            id: "F1_C14", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "ตู้เซฟเก็บพาสเวิร์ดออกตึกมีสมการ: x^2 - 5x + 6 = 0 รหัสคือค่าของ x ที่มากที่สุด รหัสคืออะไร",
            question_en: "The exit password safe shows the equation: x^2 - 5x + 6 = 0. The code is the largest value of x. What is the code?",
            answer: "3",
            explanation: "แยกตัวประกอบได้ (x-2)(x-3) = 0 ค่า x คือ 2 และ 3 ค่าที่มากที่สุดคือ 3",
            explanation_en: "Factoring gives (x-2)(x-3) = 0, so x = 2 or 3; the largest value is 3."
        },
        {
            id: "F1_C15", floor: 1, theme: "ประตูทางออกสุดท้าย", difficulty: "ยากที่สุด",
            question: "คุณมาถึงประตูสุดท้าย ป้ายเขียนว่า 'รหัสผ่านคือจำนวนชั้นทั้งหมดที่คุณผ่านมาเพื่อลงมาถึงที่นี่ คูณด้วย 10' รหัสผ่านคืออะไร",
            question_en: "You reach the final door. A sign reads: 'The password is the total number of floors you descended to get here, multiplied by 10.' What is the password?",
            answer: "50",
            explanation: "เริ่มจากชั้น 6 ลงมาถึงชั้น 1 แปลว่าผ่านมา 5 ชั้น (6->5, 5->4, 4->3, 3->2, 2->1) 5 คูณ 10 เท่ากับ 50",
            explanation_en: "Descending from Floor 6 to Floor 1 crosses 5 floors (6->5->4->3->2->1); 5 * 10 = 50."
        }
    ];

    function localizeClueObject(clue, lang) {
        if (!clue) return null;
        const copy = JSON.parse(JSON.stringify(clue));
        const activeLang = lang || (typeof global !== "undefined" && global.currentLang) || "th";
        if (activeLang === "en") {
            if (copy.question_en) copy.question = copy.question_en;
            if (copy.explanation_en) copy.explanation = copy.explanation_en;
        }
        return copy;
    }

    // ==========================================
    // API สำหรับเกม
    // ==========================================
    const EscapeTowerClueBank = {
        version: "1.1.0",
        totalClues: clueDatabase.length,
        floors: [6, 5, 4, 3, 2, 1],
        cluesPerFloor: 15,
        clueRoomsPerFloor: 3,

        /**
         * รับหมายเลขชั้น แล้วคืนค่ารายการปริศนาทั้ง 15 ข้อของชั้นนั้น
         * @param {number} floor - หมายเลขชั้น
         * @param {string} [lang] - ภาษา ('th' | 'en')
         * @returns {Array} รายการปริศนาที่ถูก Copy เพื่อป้องกันการแก้ไขต้นฉบับ
         */
        getFloorClues: function(floor, lang) {
            const numFloor = Number(floor);
            if (!this.floors.includes(numFloor)) return [];
            const clues = clueDatabase.filter(c => c.floor === numFloor);
            return clues.map(c => localizeClueObject(c, lang));
        },

        /**
         * สุ่มเลือกปริศนา 3 ข้อที่ไม่ซ้ำกันจากชั้นที่กำหนด
         * ใช้สำหรับกำหนดปริศนาให้ห้องคำใบ้ในตอนเริ่มเกม (Multiplayer Room Setup)
         * @param {number} floor - หมายเลขชั้น
         * @param {function} randomFn - ฟังก์ชันสุ่ม (default: Math.random)
         * @returns {Array} รายการปริศนา 3 ข้อ
         */
        selectCluesForGame: function(floor, randomFn = Math.random) {
            const availableClues = this.getFloorClues(floor, "th");
            if (availableClues.length === 0) return [];

            const selected = [];
            while (selected.length < this.clueRoomsPerFloor && availableClues.length > 0) {
                const randomIndex = Math.floor(randomFn() * availableClues.length);
                selected.push(availableClues[randomIndex]);
                availableClues.splice(randomIndex, 1);
            }
            return selected;
        },

        /**
         * ค้นหาข้อมูลปริศนาจาก ID
         * @param {string} clueId - รหัสปริศนา
         * @param {string} [lang] - ภาษา ('th' | 'en')
         * @returns {Object|null} ข้อมูลปริศนา หรือ null หากไม่พบ
         */
        getClueById: function(clueId, lang) {
            const clue = clueDatabase.find(c => c.id === clueId);
            return clue ? localizeClueObject(clue, lang) : null;
        },

        /**
         * คืนหมายเลขชั้นที่ปริศนานั้นอยู่
         * @param {string} clueId - รหัสปริศนา
         * @returns {number|null} หมายเลขชั้น หรือ null หากไม่พบ
         */
        getFloorForClue: function(clueId) {
            const clue = clueDatabase.find(c => c.id === clueId);
            return clue ? clue.floor : null;
        },

        /**
         * ตรวจสอบคำตอบของผู้เล่น
         * @param {string} clueId - รหัสปริศนา
         * @param {string|number} playerInput - คำตอบที่ผู้เล่นกรอก
         * @returns {boolean} true ถ้าตอบถูก, false ถ้าตอบผิดหรือไม่พบ ID
         */
        checkAnswer: function(clueId, playerInput) {
            const clue = clueDatabase.find(c => c.id === clueId);
            if (!clue || playerInput == null) return false;

            const safeAnswer = String(clue.answer).trim().toLowerCase();
            const safeInput = String(playerInput).trim().toLowerCase();

            return safeInput === safeAnswer;
        }
    };

    // เปิดเผย API สู่ Global Scope อย่างปลอดภัย
    if (typeof window !== "undefined") {
        window.EscapeTowerClueBank = EscapeTowerClueBank;
    } else if (typeof globalThis !== "undefined") {
        globalThis.EscapeTowerClueBank = EscapeTowerClueBank;
    } else {
        global.EscapeTowerClueBank = EscapeTowerClueBank;
    }

})(this);