const isAr =
    window.location.pathname.includes("/ar/") ||
    window.location.search.includes("lang=ar");

export const SourceOptions = [
    {
        key: "",
        text: isAr ? "يرجى الاختيار..." : "Please select...",
        disabled: true,
    },

    {
        key: "Alert",
        text: isAr ? "تنبيه" : "Alert",
    },
    {
        key: "AutoTicket",
        text: isAr ? "تذكرة تلقائية" : "AutoTicket",
    },
    {
        key: "Change Request",
        text: isAr ? "طلب تغيير" : "Change Request",
    },
    {
        key: "Chat",
        text: isAr ? "دردشة" : "Chat",
    },
    {
        key: "Email",
        text: isAr ? "البريد الإلكتروني" : "Email",
    },
    {
        key: "Fax",
        text: isAr ? "فاكس" : "Fax",
    },
    {
        key: "FrontRange Voice",
        text: isAr ? "نظام الصوت FrontRange" : "FrontRange Voice",
    },
    {
        key: "Incident",
        text: isAr ? "حادث" : "Incident",
    },
    {
        key: "Instant Message",
        text: isAr ? "رسالة فورية" : "Instant Message",
    },
    {
        key: "Microsoft Teams",
        text: isAr ? "مايكروسوفت تيمز" : "Microsoft Teams",
    },
    {
        key: "Network Monitor",
        text: isAr ? "مراقبة الشبكة" : "Network Monitor",
    },
    {
        key: "Phone",
        text: isAr ? "هاتف" : "Phone",
    },
    {
        key: "Self Service",
        text: isAr ? "الخدمة الذاتية" : "Self Service",
    },
    {
        key: "Tahyya",
        text: isAr ? "تحية" : "Tahyya",
    },
    {
        key: "Voice Mail",
        text: isAr ? "البريد الصوتي" : "Voice Mail",
    },
    {
        key: "Walk-in",
        text: isAr ? "حضور شخصي" : "Walk-in",
    },

];

export const UrgencyOptions = [
    {
        key: "",
        text: isAr ? "يرجى الاختيار..." : "Please select...",
        disabled: true,
    },
    {
        key: "Low",
        text: isAr ? "منخفض" : "Low",
    },
    {
        key: "Medium",
        text: isAr ? "متوسط" : "Medium",
    },
    {
        key: "High",
        text: isAr ? "مرتفع" : "High",
    },
]


export const ImpactOptions = [
    {
        key: "",
        text: isAr ? "يرجى الاختيار..." : "Please select...",
        disabled: true,
    },
    {
        key: "Low",
        text: isAr ? "منخفض" : "Low",
    },
    {
        key: "Medium",
        text: isAr ? "متوسط" : "Medium",
    },
    {
        key: "High",
        text: isAr ? "مرتفع" : "High",
    },
];

