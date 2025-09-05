import React, { useEffect, useState, useRef } from "react";
import Layout from "../../components/Layout";
import axios from "axios";
import { useNavigate, useParams } from "react-router-dom";
import "../../styles/Projects/ProjectDetails.css";
import { url, url2 } from "../../lib/api";
import Loader from "../../components/Loader";
import Swal from "sweetalert2";
import { ToastContainer } from "react-toastify";
import { MdDelete } from "react-icons/md";
import { FaEye, FaComment, FaDownload, FaCommentAlt } from "react-icons/fa";
import { toast } from "react-toastify";
import BackButton from "../../components/BackButton";
import { MdEdit } from "react-icons/md";
import Offcanvas from "../../components/OffCanvas/OffCanvas";
import { FaTelegramPlane } from "react-icons/fa";
import useRolePermissions from "../../hooks/useRolePermissions";
import InvoiceManagement from "./InvoiceManagment";
import SpinnerLoader from "../../components/SpinnerLoader";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

const ProjectDetails = () => {
  const { projectId } = useParams();
  const [project, setProject] = useState(null);
  const [loading, setLoading] = useState(true);
  const [punchListLoading, setPunchListLoading] = useState(false);
  const [allUsers, setAllUsers] = useState([]);
  const [items, setItems] = useState([]);
  const navigate = useNavigate();
  const [editableRows, setEditableRows] = useState({});
  const [punchList, setPunchList] = useState([]);
  const [showPunchModal, setShowPunchModal] = useState(false);
  const [statusMap, setStatusMap] = useState({});
  const [loadingDoc, setLoadingDoc] = useState(false);
  const [invoiceFiles, setInvoiceFiles] = useState([]);
  const [selectedUser, setSelectedUser] = useState(null);
  const [userComments, setUserComments] = useState([]);
  const [newCommentText, setNewCommentText] = useState("");
  const [isOffcanvasOpen, setIsOffcanvasOpen] = useState(false);
  const [groupedComments, setGroupedComments] = useState({});
  const [isItemCanvasOpen, setIsItemCanvasOpen] = useState(false);
  const [selectedItemId, setSelectedItemId] = useState(null);
  const [itemComments, setItemComments] = useState([]);
  const [groupedItemComments, setGroupedItemComments] = useState({});
  const [itemCommentText, setItemCommentText] = useState("");
  const [dataDoc, setDataDoc] = useState();
  const [isPunchCanvasOpen, setIsPunchCanvasOpen] = useState(false);
  const [selectedPunchItemId, setSelectedPunchItemId] = useState(null);
  const [punchComments, setPunchComments] = useState([]);
  const [groupedPunchComments, setGroupedPunchComments] = useState({});
  const [punchCommentText, setPunchCommentText] = useState("");
  const [buttonClicked, setButtonClicked] = useState(false);
  const [docsData, setDocsData] = useState({});
  const [commentLoading, setCommentLoading] = useState(false);
  const [notifyCustomerLoading, setNotifyCustomerLoading] = useState(false);
  const [currentDocType, setCurrentDocType] = useState('');
  const [isPreviewAllOpen, setIsPreviewAllOpen] = useState(false);

  const handleDownloadLeadTimePDF = () => {
    try {
      const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "A4" });

      const pageW = doc.internal.pageSize.getWidth();
      const pageH = doc.internal.pageSize.getHeight();
      const M = 40; // margins
      const title = `${project?.name || "Project"} — Lead Time Matrix`;
      const generatedAt = new Date().toLocaleString();
      const safeName = (project?.name || "project").replace(/[^\w\-]+/g, "_");

      doc.setFillColor(0, 0, 0);
      doc.rect(0, 0, pageW, 64, "F");
      doc.setFont("helvetica", "bold");
      doc.setFontSize(16);
      doc.setTextColor(255);
      doc.text(title, M, 40);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(10);
      doc.text(`Generated: ${generatedAt}`, pageW - M, 40, { align: "right" });


      const head = [["MANUFACTURER", "DESCRIPTION", "Estimated Time Of Departure", "Estimated Time Of Arrival", "ARRIVAL", "STATUS"]];
      const body = (matrix || []).map((i) => [
        i.itemName || "",
        i.quantity || "",
        i.tbdETD ? "TBD" : toDateStr(i.expectedDeliveryDate),
        i.tbdETA ? "TBD" : toDateStr(i.expectedArrivalDate),
        i.tbdArrival ? "TBD" : toDateStr(i.arrivalDate),
        (i.status || "").toUpperCase(),
      ]);

      autoTable(doc, {
        head,
        body,
        startY: 88, // directly under header
        margin: { left: M, right: M },
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 9,
          cellPadding: 6,
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
          valign: "middle",
        },
        headStyles: {
          fillColor: [0, 0, 0],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          halign: "left",
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
        columnStyles: {
          0: { cellWidth: 170 }, // Manufacturer
          1: { cellWidth: 260 }, // Description
          2: { cellWidth: 90, halign: "center" },  // ETD
          3: { cellWidth: 90, halign: "center" },  // ETA
          4: { cellWidth: 90, halign: "center" },  // Arrival
          5: { cellWidth: 90, halign: "center", fontStyle: "bold" }, // Status
        },
        didParseCell(data) {
          if (data.section === "body" && data.column.index === 5) {
            data.cell.styles.halign = "center";
          }
        },
        didDrawPage(data) {
          // Footer
          const y = pageH - 28;
          doc.setDrawColor(0);
          doc.setLineWidth(0.5);
          doc.line(M, y - 10, pageW - M, y - 10);

          doc.setFontSize(9);
          doc.setTextColor(100);
          doc.setFont("helvetica", "normal");
          doc.text("Bhouse — Lead Time Matrix", M, y);

          const pageStr = `Page ${data.pageNumber} of ${doc.getNumberOfPages()}`;
          doc.text(pageStr, pageW - M, y, { align: "right" });
        },
      });

      // ===== Summary (status counts) =====
      const finalY = doc.lastAutoTable?.finalY || 100;
      const statuses = ["Pending", "In Transit", "Delivered", "Installed", "Arrived"];
      const counts = statuses.map((s) => [
        s.toUpperCase(),
        (matrix || []).filter((r) => (r.status || "") === s).length,
      ]);

      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.setTextColor(0);
      doc.text("SUMMARY", M, finalY + 28);

      autoTable(doc, {
        head: [["STATUS", "COUNT"]],
        body: counts,
        startY: finalY + 38,
        theme: "grid",
        margin: { left: M, right: M },
        styles: {
          font: "helvetica",
          fontSize: 10,
          cellPadding: 6,
          textColor: [0, 0, 0],
          lineColor: [0, 0, 0],
          lineWidth: 0.5,
        },
        headStyles: {
          fillColor: [0, 0, 0],
          textColor: [255, 255, 255],
          fontStyle: "bold",
          halign: "left",
        },
        columnStyles: {
          0: { cellWidth: 160 },
          1: { cellWidth: 70, halign: "center" },
        },
        alternateRowStyles: { fillColor: [245, 245, 245] },
      });

      doc.save(`${safeName}_lead_time_matrix.pdf`);
    } catch (e) {
      console.error("PDF generation failed:", e);
    }
  };

  const isZeroDate = (v) =>
    v === "0000-00-00" || v === "0000-00-00 00:00:00";

  const normalizeDate = (v) => {
    if (!v || v === "TBD" || isZeroDate(v)) return null;
    return v;
  };
  const toDateStr = (d) => (!d || isZeroDate(d) ? "" : String(d).slice(0, 10));

  const handleDownloadLeadTimeExcel = () => {
    const rows = (matrix || []).map((i) => ({
      "Manufacturer Name": i.itemName || "",
      Description: i.quantity || "",
      "Estimated Time Of Departure": i.tbdETD ? "TBD" : toDateStr(i.expectedDeliveryDate),
      "Estimated Time Of Arrival": i.tbdETA ? "TBD" : toDateStr(i.expectedArrivalDate),
      Arrival: i.tbdArrival ? "TBD" : toDateStr(i.arrivalDate),
      Status: i.status || "",
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Lead Time Matrix");

    const safeName = (project?.name || "project").replace(/[^\w\-]+/g, "_");
    XLSX.writeFile(wb, `${safeName}_lead_time_matrix.xlsx`);
  };

  const scrollRef = useRef(null);
  const fileInputRef = useRef(null);
  const [activeTab, setActiveTab] = useState(
    () => localStorage.getItem("activeTab") || "overview"
  );
  const [activeTabing, setActiveTabing] = useState(
    () => localStorage.getItem("activeTabing") || "Admin"
  );
  const handleTabChange = (tabKey) => {
    setActiveTab(tabKey);
    localStorage.setItem("activeTab", tabKey);
  };

  const handleSubTabChange = (subTabKey) => {
    setActiveTabing(subTabKey);
    localStorage.setItem("activeTabing", subTabKey);
  };

  useEffect(() => {
    if (scrollRef.current) {
      scrollRef.current.scrollIntoView({ behavior: "smooth" });
    }
  }, [
    userComments,
    newCommentText,
    groupedComments,
    itemCommentText,
    punchCommentText,
    groupedPunchComments,
    groupedItemComments,
  ]);
  const docMap = {
    "Sample COI": "Sample COI",
    "COI (Certificate)": "Floor Plan ",
    "Pro Forma Invoice": "CAD file",
    // "Final Invoice": "Final Invoice",
    // "Sales Agreement": "Sales Agreement",
  };
  const normalize = (str) => str.trim().toLowerCase().replace(/\s+/g, '');

  const [data, setData] = useState();
  const fetchDocuments = async () => {
    try {
      const res = await axios.get(`${url}/customerDoc/document/${projectId}`);
      const docsArray = res.data || [];
      setData(res.data);
      console.log(docsArray, "docArray");
      const docMapData = {};
      setDataDoc(docsArray);
      docsArray.forEach((doc) => {
        docMapData[doc.documentType] = doc.filePath;
      });

      setDocsData(docMapData);
    } catch (error) {
      console.error("Failed to fetch documents", error);
    }
  };
  const refreshMatrixFromApi = async () => {
    try {
      const { data } = await axios.get(`${url}/items/${projectId}/`);
      setMatrix(data);
    } catch (e) {
      console.error("Failed to refresh preview data", e);
    }
  };

  const openPreview = async () => {
    await refreshMatrixFromApi(); // ensure fresh server data
    setIsPreviewAllOpen(true);
  };


  const fetchDocs = async () => {

    try {
      const res = await axios.get(`${url}/customerDoc/document/${projectId}`);
      setDocsData(res.data || []);
    } catch (err) {
      console.error('Failed to fetch documents:', err);
    }
  };

  const [unreadCounts, setUnreadCounts] = useState({});
  const fetchUnreadCountsForAllDocs = async () => {
    const user = JSON.parse(localStorage.getItem("user"));
    const userId = user?.user?.id;

    if (!userId) return;

    const counts = {};
    if (data) {
      await Promise?.all(
        data?.map(async (doc) => {
          try {
            const res = await axios.get(
              `${url}/customerDoc/comments/${doc.id}?userId=${userId}`
            );
            console.log(res, "res");
            const unreadComments = res.data.filter(
              (comment) => comment.User == null
            );
            const isReadFalse = unreadComments.filter(
              (comment) => comment.isRead === false
            );
            counts[doc.id] = isReadFalse.length || 0;
          } catch (err) {
            console.error(`Error fetching comments for doc ID ${doc.id}`, err);
          }
        })
      );
    }

    console.log({ counts });

    setUnreadCounts(counts);
  };
  useEffect(() => {
    fetchUnreadCountsForAllDocs();
  }, [data]);


  const [fileupload, setFileUpload] = useState(false);
  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file || !currentDocType) return;

    setLoading(true);


    const existingDoc = Object.keys(docsData).some(
      (docType) => normalize(docType) === normalize(currentDocType)
    );

    const endpoint = existingDoc ? 'update' : 'add';
    const method = existingDoc ? 'put' : 'post';

    const formData = new FormData();
    formData.append('documentType', currentDocType);
    formData.append('document', file);
    formData.append('projectId', projectId);

    try {
      const config = {
        method,
        url: `${url}/customerDoc/${endpoint}`,
        data: formData,
      };

      const res = await axios(config);
      if (res) {
        setFileUpload(true);
        // fetchDocs(); 
      }

    } catch (err) {
      console.error('Upload/Update failed:', err);
    } finally {
      setLoading(false);
      setFileUpload(false);
      setCurrentDocType('');
    }
  };


  useEffect(() => {
    fetchDocs()
    fetchDocuments()
  }, [fileupload, currentDocType]);


  const openPunchComment = async (punchId) => {
    try {
      setSelectedPunchItemId(punchId);
      const { data } = await axios.get(`${url}/punchlist/${punchId}/comments`);
      const normalized = (data || []).map((c) => {
        const isUser = c.createdByType === "user";
        const name =
          c.name ||
          c.createdByName ||
          (isUser
            ? [c.user?.firstName, c.user?.lastName].filter(Boolean).join(" ")
            : c.customer?.full_name) ||
          "Customer";

        const profileImage =
          c.profileImage ||
          (isUser ? c.user?.profileImage : c.customer?.profileImage) ||
          null;

        const userRole = c.userRole || (isUser ? c.user?.userRole : null);

        return { ...c, name, profileImage, userRole };
      });

      setPunchComments(normalized);

      const grouped = normalized.reduce((acc, comment) => {
        const date = new Date(comment.createdAt).toLocaleDateString();
        if (!acc[date]) acc[date] = [];
        acc[date].push(comment);
        return acc;
      }, {});
      setGroupedPunchComments(grouped);

      setIsPunchCanvasOpen(true);
    } catch (err) {
      console.error("Error fetching punch list comments:", err);
    }
  };

  const handleAddPunchComment = async () => {
    if (!punchCommentText.trim()) return;

    const user = JSON.parse(localStorage.getItem("user"))?.user || null;
    const customer = JSON.parse(localStorage.getItem("customer")) || null;

    const creator = user || customer;
    const creatorType = user ? "user" : "customer";

    const tempComment = {
      id: `temp-${Date.now()}`,
      comment: punchCommentText.trim(),
      createdAt: new Date().toISOString(),
      createdByType: creatorType,
      name: user
        ? [creator.firstName, creator.lastName].filter(Boolean).join(" ")
        : creator?.full_name || "Customer",
      userRole: user ? creator.userRole : null,
      profileImage: creator?.profileImage || null,
      _optimistic: true,
    };

    const today = new Date().toLocaleDateString();
    setGroupedPunchComments((prev) => ({
      ...prev,
      [today]: [...(prev[today] || []), tempComment],
    }));

    setPunchCommentText("");
    setCommentLoading(true);

    try {
      await axios.post(
        `${url}/projects/${projectId}/punchlist/${selectedPunchItemId}/comments`,
        {
          comment: tempComment.comment,
          userId: creatorType === "user" ? creator.id : null,
          clientId: creatorType === "customer" ? creator.id : null,
        }
      );
      await openPunchComment(selectedPunchItemId);
    } catch (err) {
      console.error("Failed to add punch comment:", err);
    } finally {
      setCommentLoading(false);
    }
  };


  useEffect(() => {
    fetchDocuments();
  }, [projectId]);

  const hasDate = (v) => !!normalizeDate(v);
  const validDateChoice = (dateValue, tbdFlag) => {
    const d = hasDate(dateValue);
    const t = !!tbdFlag;
    return (d && !t) || (!d && t);
  };


  const openItemComment = async (itemId) => {
    try {
      setSelectedItemId(itemId);
      const res = await axios.get(`${url}/items/${itemId}/comments`);

      const grouped = res.data.reduce((acc, comment) => {
        const date = new Date(comment.createdAt).toLocaleDateString();
        if (!acc[date]) acc[date] = [];
        acc[date].push(comment);
        return acc;
      }, {});

      setItemComments(res.data);
      setGroupedItemComments(grouped);
      setIsItemCanvasOpen(true);
    } catch (err) {
      console.error("Error fetching item comments:", err);
    }
  };

  const handleAddItemComment = async () => {
    if (!itemCommentText.trim()) return;

    const user = JSON.parse(localStorage.getItem("user"))?.user;
    const customer = JSON.parse(localStorage.getItem("customer"));
    const creator = user || customer;
    const creatorType = user ? "user" : "customer";

    const tempComment = {
      id: `temp-${Date.now()}`,
      comment: itemCommentText,
      createdAt: new Date().toISOString(),
      createdByName: creator.firstName || creator.full_name,
      userRole: user ? creator.userRole : null,
      profileImage: creator.profileImage || null,
    };

    const today = new Date().toLocaleDateString();
    setGroupedItemComments((prev) => ({
      ...prev,
      [today]: [...(prev[today] || []), tempComment],
    }));

    setItemCommentText("");
    setCommentLoading(true);

    try {
      await axios.post(`${url}/items/${selectedItemId}/comments`, {
        comment: tempComment.comment,
        createdById: creator.id,
        createdByType: creatorType,
        projectId,
      });

      await openItemComment(selectedItemId); // Sync after
    } catch (err) {
      console.error("Failed to add item comment:", err);
    } finally {
      setCommentLoading(false);
    }
  };

  const handleUploadClick = (docType) => {
    setCurrentDocType(docType);
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const formatDate = (date) => {
    const d = new Date(date);

    const options = {
      weekday: 'short',
      day: 'numeric',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    };

    // Format with lowercase am/pm
    let formatted = d.toLocaleString('en-GB', options);

    // Capitalize AM/PM
    formatted = formatted.replace(/\b(am|pm)\b/, (match) => match.toUpperCase());

    return formatted;
  };
  const [selectedFiles, setSelectedFiles] = useState({
    proposals: [],
    floorPlans: [],
    otherDocuments: [],
    presentation: [],
    salesAggrement: [],
    cad: [],

    acknowledgements: [],
    receivingReports: [],
    finalInvoice: []
  });
  const roleId = JSON.parse(localStorage.getItem("user"));
  const { rolePermissions } = useRolePermissions(roleId?.user?.roleId);

  const [newIssue, setNewIssue] = useState({
    title: "",
    issueDescription: "",
    projectItemId: "",
    productImages: [],
  });
  const [projectItems, setProjectItems] = useState([]);
  const fetchUserComments = async (toUserId) => {
    try {
      const res = await axios.get(
        `${url}/projects/${projectId}/user-comments/${toUserId}`
      );
      setUserComments(res.data);

      // Group by date
      const grouped = res.data.reduce((acc, comment) => {
        const date = new Date(comment.createdAt).toLocaleDateString();
        if (!acc[date]) acc[date] = [];
        acc[date].push(comment);
        return acc;
      }, {});
      setGroupedComments(grouped);
    } catch (err) {
      console.error("Error fetching comments:", err);
    }
  };

  const closeOffcanvas = () => {
    setIsOffcanvasOpen(false);
  };
  const fetchComments = async () => {
    if (selectedUser) {
      await fetchUserComments(selectedUser.id);
    }
  };

  const handleOpenComments = async (user) => {
    setSelectedUser(user);
    await fetchUserComments(user.id);
    setIsOffcanvasOpen(true);
  };

  const handleAddComment = async () => {
    if (!newCommentText.trim()) return;

    const storedUser = JSON.parse(localStorage.getItem("user"))?.user;
    const fromUserId = storedUser?.id;

    const tempComment = {
      id: `temp-${Date.now()}`,
      comment: newCommentText,
      createdAt: new Date().toISOString(),
      name: storedUser.firstName,
      userRole: storedUser.userRole,
      profileImage: storedUser.profileImage || null,
    };

    // Optimistic UI update
    const today = new Date().toLocaleDateString();
    setGroupedComments((prev) => ({
      ...prev,
      [today]: [...(prev[today] || []), tempComment],
    }));

    setNewCommentText("");
    setCommentLoading(true);

    try {
      await axios.post(`${url}/projects/${projectId}/user-comments`, {
        fromUserId,
        toUserId: selectedUser.id,
        comment: tempComment.comment,
        commentType: "customer",
      });

      await fetchUserComments(selectedUser.id); // Sync after
    } catch (err) {
      console.error("Error posting comment:", err);
    } finally {
      setCommentLoading(false);
    }
  };
  useEffect(() => {
    axios
      .get(`${url}/items/${projectId}`)
      .then((res) => setProjectItems(res.data))
      .catch((err) => console.error("Failed to load items:", err));
  }, [projectId, showPunchModal]);
  const fetchPunchList = async () => {
    try {
      const res = await axios.get(`${url}/projects/${projectId}/punch-list`);
      const parsed = res.data.map((issue) => ({
        ...issue,
        productImages:
          typeof issue.productImages === "string"
            ? JSON.parse(issue.productImages)
            : Array.isArray(issue.productImages)
              ? issue.productImages
              : [],
      }));
      setPunchList(parsed);
      const initialStatus = {};
      parsed.forEach((item) => {
        initialStatus[item.id] = item.status || "Pending";
      });
      setStatusMap(initialStatus);
    } catch (err) {
      console.error("Error fetching punch list:", err);
    }
  };
  useEffect(() => {
    fetchPunchList();
  }, [projectId, showPunchModal]);

  const handleItemChange = (index, field, value) => {
    const updated = [...items];
    updated[index][field] = value;
    setItems(updated);
  };

  const updateItem = async (item) => {
    try {
      const { data } = await axios.put(`${url}/items/project-items/${item.id}`, item);
      toast.success("Item updated!");
      const saved = data || item;
      setMatrix((prev) => prev.map((it) => (it.id === saved.id ? saved : it)));
      setItems((prev) => prev.map((it) => (it.id === saved.id ? saved : it)));
    } catch (err) {
      toast.error("Error updating item.");
      console.error(err);
    }
  };

  const deleteItem = async (id) => {
    try {
      await axios.delete(`${url}/items/project-items/${id}`);
      setItems((prev) => prev.filter((item) => item.id !== id));
      setMatrix((prev) => prev.filter((item) => item.id !== id));
      toast.success("Item deleted!");
    } catch (err) {
      toast.error("Error deleting item.");
      console.error(err);
    }
  };

  const addNewItemToBackend = async (item, index) => {
    setButtonClicked(true);
    try {
      const res = await axios.post(`${url}/items/project-items`, item);
      const updated = [...items];
      updated[index] = res.data;
      setItems(updated);
      setMatrix(updated);

      toast.success("Item added!");
    } catch (err) {
      alert("Failed to add item.");
      console.error(err);
    } finally {
      setButtonClicked(false);
    }
  };

  const handleAddItemRow = () => {
    setItems((prev) => [
      ...prev,
      {
        itemName: "",
        quantity: "",
        expectedDeliveryDate: null,
        expectedArrivalDate: null,
        status: "Pending",
        projectId,
        arrivalDate: null
      },
    ]);
  };
  const [read, unread] = useState();
  const fetchVisibleUserComments = async (users) => {
    try {
      const commentCounts = await Promise.all(
        users.map(async (user) => {
          const { data } = await axios.get(
            `${url}/projects/${projectId}/user-comments/${user.id}`
          );

          const unreadComments = data.filter(
            (comment) => comment.createdByType === "customer"
          );
          const filterIsReadFalse = unreadComments.filter(
            (comment) => comment.isRead === false
          );

          return { id: user.id, commentCount: filterIsReadFalse.length };
        })
      );

      // Do something with commentCounts if needed
      // setVisibleUserComments(commentCounts);
      unread(commentCounts);
    } catch (err) {
      console.error("Error fetching visible user comment counts", err);
    }
  };
  const markCommentsAsRead = async (toUserId) => {
    try {
      const response = await axios.put(
        `${url}/projects/${projectId}/teamMarkCommentsAsRead/${toUserId}`
      );
      unread();
    } catch (error) {
      console.log(error);
    }
  };
  useEffect(() => {
    if (project && project.assignedTeamRoles && allUsers.length > 0) {
      const extractedUsers = [];

      project?.assignedTeamRoles?.forEach((roleGroup) => {
        roleGroup.users.forEach((userId) => {
          const user = allUsers.find(
            (u) => u.id.toString() === userId.toString()
          );
          if (user) {
            extractedUsers.push(user);
          }
        });
      });

      fetchVisibleUserComments(extractedUsers);
    }
  }, [allUsers]);

  const [commentCountsByManufacturerId, setCommentCountsByManufacturerId] =
    useState({});
  const fetchItemsComments = async () => {
    try {
      const commentCounts = {};

      for (const manuId in itemsByManufacturerId) {
        const itemIds = itemsByManufacturerId[manuId].map((item) => item.id);

        const commentPromises = itemIds.map((id) =>
          axios.get(`${url}/items/${id}/comments`).catch(() => ({ data: [] }))
        );

        const results = await Promise.all(commentPromises);
        const allComments = results.flatMap((res) => res.data || []);
        const userComments = allComments.filter(
          (cmt) => cmt.createdByType === "customer"
        );
        const isReadFalse = userComments.filter((item) => item.isRead == false);
        commentCounts[manuId] = isReadFalse.length;
      }
      setCommentCountsByManufacturerId(commentCounts);
    } catch (error) {
      console.log("Error fetching comments:", error);
    }
  };
  useEffect(() => {
    fetchItemsComments();
  }, [items]);
  const [matrix, setMatrix] = useState([]);
  const [itemsByManufacturerId, setItemsByManufacturerId] = useState({});
  useEffect(() => {
    const fetchItems = async () => {
      try {
        const res = await axios.get(`${url}/items/${projectId}/`);
        const grouped = res.data.reduce((acc, item) => {
          const manufacturer = item.id || "Unknown";
          if (!acc[manufacturer]) acc[manufacturer] = [];
          acc[manufacturer].push(item);
          return acc;
        }, {});
        setItemsByManufacturerId(grouped);
        setItems(res.data);
        setMatrix(res.data);
      } catch (error) {
        console.error("Error fetching items:", error);
      }
    };

    fetchItems();
  }, [projectId]);

  useEffect(() => {
    const fetchProjectDetails = async () => {
      try {
        const [projectRes, usersRes] = await Promise.all([
          axios.get(`${url}/projects/${projectId}`),
          axios.get(`${url}/auth/getAllUsers`),
        ]);

        const fetchedProject = projectRes.data;
        const fetchedUsers = usersRes.data;

        fetchedProject.assignedTeamRoles = Array.isArray(
          fetchedProject.assignedTeamRoles
        )
          ? fetchedProject.assignedTeamRoles
          : JSON.parse(fetchedProject.assignedTeamRoles || "[]");

        fetchedProject.proposals = Array.isArray(fetchedProject.proposals)
          ? fetchedProject.proposals
          : JSON.parse(fetchedProject.proposals || "[]");

        fetchedProject.floorPlans = Array.isArray(fetchedProject.floorPlans)
          ? fetchedProject.floorPlans
          : JSON.parse(fetchedProject.floorPlans || "[]");

        fetchedProject.otherDocuments = Array.isArray(
          fetchedProject.otherDocuments
        )
          ? fetchedProject.otherDocuments
          : JSON.parse(fetchedProject.otherDocuments || "[]");
        fetchedProject.invoice = Array.isArray(fetchedProject.invoice)
          ? fetchedProject.invoice
          : JSON.parse(fetchedProject.invoice || "[]");
        fetchedProject.cad = Array.isArray(fetchedProject.cad)
          ? fetchedProject.cad
          : JSON.parse(fetchedProject.cad || "[]");
        fetchedProject.salesAggrement = Array.isArray(
          fetchedProject.salesAggrement
        )
          ? fetchedProject.salesAggrement
          : JSON.parse(fetchedProject.salesAggrement || "[]");
        fetchedProject.presentation = Array.isArray(fetchedProject.presentation)
          ? fetchedProject.presentation
          : JSON.parse(fetchedProject.presentation || "[]");

        fetchedProject.acknowledgements = Array.isArray(
          fetchedProject.acknowledgements
        )
          ? fetchedProject.acknowledgements
          : JSON.parse(fetchedProject.acknowledgements || "[]");

        fetchedProject.finalInvoice = Array.isArray(
          fetchedProject.finalInvoice
        )
          ? fetchedProject.finalInvoice
          : JSON.parse(fetchedProject.finalInvoice || "[]");




        setInvoiceFiles(fetchedProject.invoice);

        setProject(fetchedProject);
        setAllUsers(fetchedUsers);
        setLoading(false);
      } catch (error) {
        console.error("Error fetching project details", error);
      }
    };

    fetchProjectDetails();
  }, [projectId]);

  const [commentCountsByIssueId, setCommentCountsByIssueId] = useState({});
  const fetchCommentsForIssues = async () => {
    const counts = {};
    console.log({ punchList });
    await Promise.all(
      punchList.map(async (issue) => {
        try {
          const res = await axios.get(`${url}/punchlist/${issue.id}/comments`);
          const unreadUserComments = res.data.filter(
            (comment) =>
              comment.isRead === false && comment.createdByType == "customer"
          );
          counts[issue.id] = unreadUserComments.length;
        } catch (err) {
          console.error(`Error fetching comments for issue ${issue.id}:`, err);
          counts[issue.id] = 0;
        }
      })
    );

    setCommentCountsByIssueId(counts);
  };
  const markPunchListItemCommentsAsRead = async (punchListItemId) => {
    try {
      const response = await axios.put(
        `${url}/projects/markPunchListItemCommentsAsRead/${punchListItemId}`
      );
      if (response) {
        setCommentCountsByIssueId({});
      }
    } catch (error) {
      console.log(error);
    }
  };

  const [commentCounts, setCommentCounts] = useState({});
  const fetchAllComments = async () => {
    const allFileFields = [
      "proposals",
      "floorPlans",
      "cad",
      "salesAggrement",
      "presentation",
      "otherDocuments",
      "acknowledgements",
      "receivingReports",
      "finalInvoive"
    ];

    const newCommentCounts = {};
    if (allFileFields) {
      var res = await axios.get(`${url}/projects/${projectId}`);
      let project = res.data;
      console.log({ project });
      for (const field of allFileFields) {
        const files = JSON.parse(project[field] || "[]");

        for (const file of files) {
          let filePath = file.url || file.filePath || file; // adjust based on your file object structure

          if (filePath.startsWith("/")) {
            filePath = filePath.substring(1);
          }
          console.log(filePath);

          try {
            const res = await axios.get(
              `${url}/projects/${projectId}/file-comments`,
              {
                params: { filePath },
              }
            );
            console.log(res, "res");

            // Filter comments where isRead is false
            const unreadComments = res.data.filter(
              (comment) => comment.isRead === false && comment.user == null
            );
            console.log({ unreadComments });
            newCommentCounts[filePath] = unreadComments.length || 0;
          } catch (err) {
            console.error(`Failed to fetch comments for ${filePath}:`, err);
            newCommentCounts[filePath] = 0;
          }
        }
      }
    }

    console.log({ newCommentCounts });
    setCommentCounts(newCommentCounts);
  };
  const documentMarkCommentsAsRead = async (filePath) => {
    try {
      const res = await axios.put(
        `${url}/documentMarkCommentsAsRead`,
        {},
        {
          params: { filePath },
        }
      );

      setCommentCounts({});
    } catch (error) {
      console.log("Error updating comments:", error);
    }
  };

  useEffect(() => {
    fetchAllComments();
  }, [project, projectId]);

  useEffect(() => {
    fetchCommentsForIssues();
  }, [punchList]);

  if (loading) {
    return (
      <Layout>
        <div className="project-details-header">
          <h1>Project Details</h1>
        </div>
        <div className="loading">
          <Loader />
        </div>
      </Layout>
    );
  }
  const removeRow = (index) => {
    setItems((prev) => prev.filter((_, i) => i !== index));
    setMatrix((prev) => prev.filter((_, i) => i !== index));
  };


  // const handleFileUpload = (e, category) => {
  //   const files = Array.from(e.target.files);

  //   setSelectedFiles((prev) => ({
  //     ...prev,
  //     [category]: files,
  //   }));
  // };



  const handleFileUpload = (e, category) => {
    const files = Array.from(e.target.files);
    setSelectedFiles((prev) => ({
      ...prev,
      [category]: [...(prev[category] || []), ...files],
    }));
  };

  const removeSelectedFile = (category, index) => {
    setSelectedFiles((prev) => ({
      ...prev,
      [category]: prev[category].filter((_, i) => i !== index),
    }));
  };

  const uploadSelectedFiles = async (category) => {
    if (!selectedFiles[category].length) return;

    const formData = new FormData();
    selectedFiles[category].forEach((file) => formData.append("files", file));
    formData.append("category", category);

    try {
      setLoadingDoc(true);
      await axios
        .post(
          `${url}/projects/${projectId}/upload-files?category=${category}`,
          formData
        )
        .then((res) => {
          if (res) {
            setLoadingDoc(false);
          }
        });

      // Clear selected files
      setSelectedFiles((prev) => ({ ...prev, [category]: [] }));

      // Refresh project data
      const res = await axios.get(`${url}/projects/${projectId}`);
      const updatedProject = res.data;
      [
        "proposals",
        "floorPlans",
        "otherDocuments",
        "presentation",
        "salesAggrement",
        "cad",
        "acknowledgements",
        "receivingReports",
        "finalInvoice"
      ].forEach((key) => {
        updatedProject[key] = Array.isArray(updatedProject[key])
          ? updatedProject[key]
          : JSON.parse(updatedProject[key] || "[]");
      });

      setProject(updatedProject);
    } catch (error) {
      console.error("Upload error:", error);
      toast.error("Upload failed.");
      setLoadingDoc(false);
    }
  };

  const parseSafeArray = (value) => {
    if (Array.isArray(value)) return value;
    try {
      return JSON.parse(value) || [];
    } catch {
      return [];
    }
  };

  const handleProjectFileUpdate = async (filePath, category) => {
    const existingFiles = parseSafeArray(project[category]);

    const updatedCategoryFiles = existingFiles.filter((f) => f !== filePath);
    const updatedProject = {
      ...project,
      [category]: updatedCategoryFiles,
    };

    const formDataToSend = new FormData();
    formDataToSend.append("removedFiles", JSON.stringify([filePath]));

    // Append required fields
    [
      "name",
      "type",
      "clientName",
      "description",
      "startDate",
      "estimatedCompletion",
      "totalValue",
      "deliveryAddress",
      "deliveryHours",
    ].forEach((key) => {
      if (updatedProject[key] !== undefined) {
        formDataToSend.append(key, updatedProject[key]);
      }
    });

    try {
      const res = await fetch(`${url}/projects/${projectId}`, {
        method: "PUT",
        body: formDataToSend,
      });

      if (res.status === 200) {
        setProject((prev) => ({
          ...prev,
          [category]: updatedCategoryFiles,
        }));
        toast.success("File removed and project updated!");
      } else {
        const data = await res.json();
        toast.error(data?.error || "Failed to update project.");
      }
    } catch (error) {
      console.error("Update failed", error);
      toast.error("Error while updating project.");
    }
  };

  const handleCustomerFileDelete = async (docType) => {
    try {
      const res = await axios.delete(`${url}/customerDoc/delete`, {
        data: {
          documentType: docType,
          projectId: projectId
        }
      });

      if (res?.status === 200 || res?.status === 204) {
        fetchDocs(); // refresh view
        Swal.fire("Deleted!", "File has been removed.", "success");
      }
    } catch (err) {
      console.error("Delete failed:", err);
      Swal.fire("Error", "File deletion failed. Try again.", "error");
    }
  };

  const toggleEditRow = (index) => {
    setEditableRows((prev) => ({
      ...prev,
      [index]: !prev[index],
    }));
  };

  const handlePunchStatusUpdate = async (id, status) => {
    try {
      await axios.put(`${url}/punch-list/${id}/status`, { status });
      toast.success("Punch List status updated!");

      // Refresh punch list after update
      const res = await axios.get(`${url}/projects/${projectId}/punch-list`);
      const parsed = res.data.map((issue) => ({
        ...issue,
        productImages:
          typeof issue.productImages === "string"
            ? JSON.parse(issue.productImages)
            : Array.isArray(issue.productImages)
              ? issue.productImages
              : [],
      }));
      setPunchList(parsed);
    } catch (err) {
      console.error("Status update failed:", err);
      toast.error("Failed to update punch list status.");
    }
  };
  const handleImageSelect = (e) => {
    const files = Array.from(e.target.files);
    const totalFiles = [...newIssue.productImages, ...files];

    if (totalFiles.length > 5) {
      toast.error("You can only upload up to 5 images.");
      return;
    }

    setNewIssue((prev) => ({
      ...prev,
      productImages: totalFiles,
    }));
  };
  const removeImage = (index) => {
    setNewIssue((prev) => ({
      ...prev,
      productImages: prev.productImages.filter((_, i) => i !== index),
    }));
  };

  const tabModules = [
    {
      key: "overview",
      label: "Overview",
      permissionKey: null,
      alwaysVisible: true,
    },
    { key: "documents", label: "Documents", permissionKey: "ProjectDocument" },
    { key: "team", label: "Team", permissionKey: "AssignedTeamComments" },
    {
      key: "leadTimeMatrix",
      label: "Lead Time Matrix",
      permissionKey: null,
      alwaysVisible: true,
    },
    { key: "punchlist", label: "Punch List", permissionKey: "PunchList" },
    { key: "invoice", label: "Invoice", permissionKey: "Invoicing" },
    // { key: "settings", label: "Settings", permissionKey: null, alwaysVisible: true  },
  ];
  // notifyTOCustomer
  const handleToNotifyCustomer = async () => {
    const data = { projectId };

    Swal.fire({
      title: "Are you sure?",
      text: "Do you want to notify the customer?",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#d33",
      cancelButtonColor: "#3085d6",
      confirmButtonText: "Yes, Notify!",
    }).then(async (result) => {
      if (result.isConfirmed) {
        try {
          setNotifyCustomerLoading(true);
          const response = await axios.post(
            `${url}/notifyCustomerOfItemStatus`,
            data
          );
          console.log(response);
          setNotifyCustomerLoading(false);
          await Swal.fire({
            title: "Success!",
            text: "Customer has been notified successfully.",
            icon: "success",
            timer: 2000, // 3 seconds
            showConfirmButton: false,
          });
        } catch (error) {
          console.error(error);
          setNotifyCustomerLoading(false);
          Swal.fire("Error!", "Something went wrong while notifying.", "error");
        }
      }
    });
  };

  const getCommentCountByUserId = (userId) => {
    const entry = read?.find((c) => c.id === userId);
    return entry ? entry.commentCount : 0;
  };
  const itemMarkItemCommentsAsRead = async (itemId) => {
    console.log(itemId, "itemId");
    try {
      const response = await axios.put(
        `${url}/projects/itemMarkItemCommentsAsRead/${itemId}`
      );
      if (response) {
        setCommentCountsByManufacturerId({});
      }
    } catch (error) {
      console.log(error);
    }
  };

  const markReadCustomerDocComment = async (id) => {
    let res = await axios.put(
      `${url}/customerDoc/updateCommentsIsReadByDocumentId/${id}`
    );
    if (res) {
      setUnreadCounts({});
    }
  };

  const handleToNotifyCustomerofPunchList = async () => {
    const data = { projectId };

    Swal.fire({
      title: "Are you sure?",
      text: "Do you want to notify the customer?",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#d33",
      cancelButtonColor: "#3085d6",
      confirmButtonText: "Yes, Notify!",
    }).then(async (result) => {
      if (result.isConfirmed) {
        try {
          setNotifyCustomerLoading(true);
          const response = await axios.post(
            `${url}/notifyCustomerOfPunchList`,
            data
          );
          console.log(response);
          setNotifyCustomerLoading(false);
          await Swal.fire({
            title: "Success!",
            text: "Customer has been notified successfully.",
            icon: "success",
            timer: 2000, // 3 seconds
            showConfirmButton: false,
          });
        } catch (error) {
          console.error(error);
          setNotifyCustomerLoading(false);
          Swal.fire("Error!", "Something went wrong while notifying.", "error");
        }
      }
    });
  };
  const parseArrayish = (v) => {
    if (Array.isArray(v)) return v;
    if (typeof v === "string") {
      try {
        const parsed = JSON.parse(v);
        if (Array.isArray(parsed)) return parsed;
      } catch (_) { }
      return v ? v.split(",").map(s => s.trim()).filter(Boolean) : [];
    }
    if (v && typeof v === "object") {
      const maybe = v.name || [v.firstName, v.lastName].filter(Boolean).join(" ");
      return maybe ? [maybe] : [];
    }
    return [];
  };

  const formatClientName = (value) => {
    const arr = parseArrayish(value);
    return arr.join(", ");
  };

  return (
    <Layout>
      <ToastContainer />
      <div className="project-details-page">
        <BackButton />
        <div className="project-details-header">
          <h1>{project.name}</h1>
          <p className="project-subtitle">{project.type} Project</p>
        </div>

        {loadingDoc ? (
          <div className="doc-loader">
            <Loader />
          </div>
        ) : (
          <>
            <div className="tabs">
              {tabModules.map(
                ({ key, label, permissionKey, alwaysVisible }) => {
                  const hasPermission =
                    alwaysVisible || // Show if marked alwaysVisible
                    (permissionKey && rolePermissions?.[permissionKey]?.view);
                  return hasPermission ? (
                    <button
                      key={key}
                      className={activeTab === key ? "tab active" : "tab"}
                      onClick={() => handleTabChange(key)}
                    >
                      {label}
                    </button>
                  ) : null;
                }
              )}
            </div><input
              type="file"
              ref={fileInputRef}
              onChange={handleFileChange}
              style={{ display: "none" }}
            />
            {/* Project Tabing View */}
            <div className="tab-content">
              {activeTab === "overview" && (
                <div className="project-details-container">
                  <div className="project-info-card">
                    <h2>Project Overview</h2>
                    <div className="info-group">
                      <strong>Client:</strong> {formatClientName(project.clientName)}
                    </div>
                    <div className="info-group">
                      <strong>Status:</strong> {project.status}
                    </div>
                    <div className="info-group">
                      <strong>Type:</strong> {project.type}
                    </div>
                    <div className="info-group">
                      <strong>Description:</strong>{" "}
                      {project.description || "N/A"}
                    </div>

                    <div className="info-group">
                      <strong>Estimated Occupancy Date:</strong>
                      {project.estimatedCompletion}
                    </div>
                    <div className="info-group">
                      <strong>Total Value:</strong> ${" "}
                      {project.totalValue?.toLocaleString() || "N/A"}
                    </div>
                    <div className="info-group">
                      <strong>Advance Payment:</strong> ${" "}
                      {project.advancePayment?.toLocaleString() || "N/A"}
                    </div>
                  </div>
                  <div className="project-info-card">
                    <h2>Delivery Details</h2>
                    <div className="info-group">
                      <strong>Address:</strong>{" "}
                      {project.deliveryAddress || "N/A"}
                    </div>
                    <div className="info-group">
                      <strong>Hours:</strong> {project.deliveryHours || "N/A"}
                    </div>
                    <div className="info-group">
                      <strong>Delivery Date:</strong> {project.deliveryDate || "N/A"}
                    </div>

                  </div>

                  <div className="project-info-card">
                    <h2>Point of Contact</h2>
                    <div className="info-group">
                      <strong>Name:</strong>{" "}
                      {project.pocName === "null" ? "N/A" : project.pocName || "NA"}
                    </div>
                    <div className="info-group">
                      <strong>Email:</strong>  {project.pocEmail === "null" ? "N/A" : project.pocEmail || "NA"}
                    </div>
                    <div className="info-group">
                      <strong> Phone</strong> {project.pocNumber === "null" ? "N/A" : project.pocNumber || "NA"}
                    </div>
                  </div>
                </div>
              )}
              {activeTab === "documents" && (
                <div className="project-info-card">
                  <div className="tabs-container">
                    <div className="tabs-header">
                      <button
                        className={`tab-button ${activeTabing === "Admin" ? "active" : ""
                          }`}
                        onClick={() => handleSubTabChange("Admin")}
                      >
                        BHOUSE
                      </button>
                      <button
                        className={`tab-button ${activeTabing === "Customer" ? "active" : ""
                          }`}
                        onClick={() => handleSubTabChange("Customer")}
                      >
                        CUSTOMER
                      </button>
                    </div>
                  </div>
                  <div className="tab-content">
                    {activeTabing === "Admin" && (
                      <div className="tab-panel">
                        <h2>Uploaded Documents</h2>
                        {[
                          {
                            title: "Detailed Proposal",
                            files: project.proposals,
                            category: "proposals",
                          },
                          {
                            title: "Pro Forma Invoice",
                            files: project.floorPlans,
                            category: "floorPlans",
                          },
                          {
                            title: "Product Maintenance",
                            files: project.otherDocuments,
                            category: "otherDocuments",
                          },
                          {
                            title: " COI(Certificate)",
                            files: project.cad,
                            category: "cad",
                          },
                          {
                            title: "Options Presentation",
                            files: project.presentation,
                            category: "presentation",
                          },
                          {
                            title: "Sales Agreement",
                            files: project.salesAggrement,
                            category: "salesAggrement",
                          },
                          {
                            title: "Acknowledgements",
                            files: project.acknowledgements,
                            category: "acknowledgements",
                          },
                          {
                            title: "Receiving Reports",
                            files: project.receivingReports,
                            category: "receivingReports",
                          },
                          {
                            title: "Final Invoice",
                            files: project.finalInvoice,
                            category: "finalInvoice",
                          },
                        ].map((docCategory, idx) => (
                          <div key={idx} className="  -section">
                            <h3>{docCategory.title.toUpperCase()}</h3>
                            {rolePermissions?.ProjectDocument?.add ? (
                              <input
                                type="file"
                                disabled={false}
                                multiple
                                onChange={(e) =>
                                  handleFileUpload(e, docCategory.category)
                                }
                              />
                            ) : null}

                            {selectedFiles[docCategory.category]?.length >
                              0 && (
                                <div className="file-preview-section">
                                  <h4>Files to be uploaded:</h4>
                                  <ul className="preview-list">
                                    {selectedFiles[docCategory.category].map(
                                      (file, i) => (
                                        <li key={i} className="preview-item">
                                          {file.name}
                                          <span
                                            className="remove-preview"
                                            onClick={() =>
                                              removeSelectedFile(
                                                docCategory.category,
                                                i
                                              )
                                            }
                                          >
                                            ×
                                          </span>
                                        </li>
                                      )
                                    )}
                                  </ul>
                                  <button
                                    className="upload-btn"
                                    onClick={() =>
                                      uploadSelectedFiles(docCategory.category)
                                    }
                                  >
                                    Upload
                                  </button>
                                </div>
                              )}
                            {(() => {
                              const files = Array.isArray(docCategory?.files)
                                ? docCategory.files
                                : typeof docCategory.files === "string"
                                  ? JSON.parse(docCategory.files)
                                  : [];

                              return files.length > 0 ? (
                                <div className="uploaded-files">
                                  {files.map((filePath, idx) => {
                                    const fileName = filePath?.split("/").pop();
                                    const fileUrl = filePath.startsWith(
                                      "uploads"
                                    )
                                      ? `${url2}/${filePath}`
                                      : filePath;

                                    const handleDownload = async () => {
                                      try {
                                        const response = await fetch(fileUrl);
                                        const blob = await response.blob();
                                        const downloadUrl =
                                          window.URL.createObjectURL(blob);
                                        const link =
                                          document.createElement("a");
                                        link.href = downloadUrl;
                                        link.download = fileName;
                                        document.body.appendChild(link);
                                        link.click();
                                        link.remove();
                                        window.URL.revokeObjectURL(downloadUrl);
                                      } catch (error) {
                                        console.error("Download failed", error);
                                        alert(
                                          "Download failed, please try again."
                                        );
                                      }
                                    };

                                    return (
                                      <div
                                        key={idx}
                                        className="file-item-enhanced"
                                      >
                                        <span className="file-name-enhanced">
                                          {fileName}
                                        </span>
                                        <div className="file-actions">
                                          {![".dwg", ".dxf", ".cad"].includes(
                                            fileUrl
                                              ?.split(".")
                                              .pop()
                                              ?.toLowerCase()
                                              .startsWith(".")
                                              ? fileUrl
                                                .split(".")
                                                .pop()
                                                .toLowerCase()
                                              : `.${fileUrl
                                                .split(".")
                                                .pop()
                                                .toLowerCase()}`
                                          ) && (
                                              <button
                                                className="file-action-btn"
                                                onClick={() =>
                                                  window.open(fileUrl, "_blank")
                                                }
                                                title="View"
                                              >
                                                <FaEye />
                                              </button>
                                            )}
                                          <button
                                            className="file-action-btn"
                                            onClick={handleDownload}
                                            title="Download"
                                          >
                                            <FaDownload />
                                          </button>
                                          <button
                                            className="file-action-btn"
                                            title="Delete"
                                            onClick={() => {
                                              Swal.fire({
                                                title: "Are you sure?",
                                                text: "Do you want to remove this file?",
                                                icon: "warning",
                                                showCancelButton: true,
                                                confirmButtonText:
                                                  "Yes, remove it!",
                                                cancelButtonText: "Cancel",
                                              }).then(async (result) => {
                                                if (result.isConfirmed) {
                                                  await handleProjectFileUpdate(
                                                    filePath,
                                                    docCategory.category
                                                  );
                                                }
                                              });
                                            }}
                                          >
                                            <MdDelete />
                                          </button>
                                          <div
                                            onClick={() =>
                                              documentMarkCommentsAsRead(
                                                filePath
                                              )
                                            }
                                          >
                                            <button
                                              className="file-action-btn"
                                              title="Comment"
                                              onClick={() =>
                                                navigate(
                                                  `/project/${projectId}/file-comments`,
                                                  {
                                                    state: {
                                                      filePath,
                                                      category:
                                                        docCategory.category,
                                                    },
                                                  }
                                                )
                                              }
                                            >
                                              <FaComment />
                                              {commentCounts[filePath] > 0 && (
                                                <span
                                                  style={{
                                                    color: "red",
                                                    fontWeight: "bold",
                                                  }}
                                                >
                                                  {" "}
                                                  ({commentCounts[filePath]})
                                                </span>
                                              )}
                                            </button>
                                          </div>
                                        </div>
                                      </div>
                                    );
                                  })}
                                </div>
                              ) : (
                                <p>No documents uploaded.</p>
                              );
                            })()}
                          </div>
                        ))}
                      </div>
                    )}


                    {activeTabing === "Customer" && (
                      <div className="tab-panel">
                        <h2>Uploaded Documents</h2>

                        {Object.keys(docMap).map((key, idx) => {
                          const normalizedKey = normalize(key);
                          const fileEntry = Object.entries(docsData).find(
                            ([docType]) => normalize(docType) === normalizedKey
                          );
                          const filePath = fileEntry?.[1];
                          const fileName = filePath?.split("/").pop();
                          const fileUrl = filePath?.startsWith("uploads")
                            ? `${url2}/${filePath}`
                            : filePath;

                          const matchedDoc = dataDoc?.find(
                            (doc) => normalize(doc.documentType) === normalizedKey
                          );
                          const documentId = matchedDoc?.id;
                          const uploaderName =
                            (matchedDoc?.uploadedByCustomerName && matchedDoc.uploadedByCustomerName.trim()) ||
                            "Bhouse";

                          const uploadedAt = matchedDoc?.updatedAt || matchedDoc?.createdAt || null;


                          const docType = fileEntry?.[0] || key;

                          return (
                            <div key={idx} className="doc-view-section">
                              <h4>{docMap[key].toUpperCase()}</h4>



                              {filePath ? (
                                <div className="file-item-enhanced">
                                  <div>
                                    <span className="file-name-enhanced">{fileName}</span>

                                    {/* Always show; default to Bhouse if no name */}
                                    <div className="file-meta" style={{ marginTop: 4, fontSize: 12, color: "#666" }}>
                                      Uploaded by <strong>{uploaderName}</strong>
                                      {uploadedAt && (
                                        <>
                                          {" "}•{" "}
                                          {new Date(uploadedAt).toLocaleString()}
                                        </>
                                      )}
                                    </div>
                                  </div>


                                  <div className="uploaded-icon">
                                    <button
                                      className="file-action-btn eye"
                                      onClick={() => window.open(fileUrl, "_blank")}
                                      title="View"
                                    >
                                      <FaEye />
                                    </button>

                                    <button
                                      className="file-action-btn"
                                      title="Download"
                                      onClick={async () => {
                                        try {
                                          const response = await fetch(fileUrl);
                                          const blob = await response.blob();
                                          const downloadUrl = window.URL.createObjectURL(blob);
                                          const a = document.createElement("a");
                                          a.href = downloadUrl;
                                          a.download = fileName;
                                          document.body.appendChild(a);
                                          a.click();
                                          a.remove();
                                          window.URL.revokeObjectURL(downloadUrl);
                                        } catch (error) {
                                          console.error("Download failed", error);
                                          alert("Download failed, please try again.");
                                        }
                                      }}
                                    >
                                      <FaDownload />
                                    </button>

                                    <button
                                      className="file-action-btn"
                                      title="Delete"
                                      onClick={() => {
                                        Swal.fire({
                                          title: "Are you sure?",
                                          text: "Do you want to delete this document?",
                                          icon: "warning",
                                          showCancelButton: true,
                                          confirmButtonText: "Yes, delete it!",
                                          cancelButtonText: "Cancel",
                                        }).then(async (result) => {
                                          if (result.isConfirmed) {
                                            await handleCustomerFileDelete(docType);
                                          }
                                        });
                                      }}
                                    >
                                      <MdDelete />
                                    </button>

                                    {matchedDoc && (
                                      <div onClick={() => markReadCustomerDocComment(documentId)}>
                                        <button
                                          className="file-action-btn"
                                          onClick={() =>
                                            navigate(`/customerDoc/comment/${docType}/${documentId}`, {
                                              state: {
                                                data: dataDoc,
                                                fileName: docType,
                                                filePath,
                                              },
                                            })
                                          }
                                          title="Comment"
                                        >
                                          <FaComment />
                                        </button>

                                        {unreadCounts[documentId] > 0 && (
                                          <span style={{ color: "red", fontWeight: "bold" }}>
                                            ({unreadCounts[documentId]})
                                          </span>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                </div>
                              ) : (
                                <>
                                  <p>No document uploaded.</p>
                                  <div style={{ marginTop: "8px" }}>
                                    <button onClick={() => handleUploadClick(docType)}>
                                      {filePath ? "Update" : "Upload"}
                                    </button>
                                  </div>
                                </>
                              )}




                            </div>
                          );
                        })}
                      </div>
                    )}

                  </div>
                </div>
              )}
              {activeTab === "team" && (
                <div className="project-info-card">
                  <h2>Assigned Team</h2>
                  {Array.isArray(project.assignedTeamRoles) && project.assignedTeamRoles.length > 0 ? (
                    <div className="team-grid">
                      {project.assignedTeamRoles.map((roleGroup, index) => (
                        <div key={index} className="role-card">
                          <h3 className="role-title">{roleGroup.role}</h3>
                          {roleGroup.users.map((userId) => {
                            const user = allUsers.find(
                              (u) => u.id.toString() === userId.toString()
                            );
                            return user ? (
                              <div
                                key={user.id}
                                className="user-card-horizontal"
                              >
                                <img
                                  src={
                                    user.profileImage
                                      ? `${url2}/${user.profileImage}`
                                      : `${process.env.PUBLIC_URL}/assets/Default_pfp.jpg`
                                  }
                                  alt={`${user.firstName} ${user.lastName}`}
                                  className="user-profile-img-horizontal"
                                />
                                <div className="user-info-horizontal">
                                  <span className="user-name-horizontal">
                                    {user.firstName} {user.lastName}
                                  </span>
                                  <div
                                    onClick={() => markCommentsAsRead(user.id)}
                                  >
                                    <button
                                      className="comment-btna"
                                      onClick={() => handleOpenComments(user)}
                                    >
                                      <FaCommentAlt />
                                      {getCommentCountByUserId(user.id) > 0 && (
                                        <span
                                          style={{
                                            fontSize: "20px",
                                            color: "red",
                                            padding: "2px",
                                          }}
                                        >
                                          ({getCommentCountByUserId(user.id)})
                                        </span>
                                      )}
                                    </button>
                                  </div>
                                </div>
                              </div>
                            ) : null;
                          })}
                        </div>
                      ))}
                    </div>
                  ) : (
                    <p>No team assigned.</p>
                  )}
                </div>
              )}
              {activeTab === "leadTimeMatrix" && (
                <div className="project-info-card">
                  <div className="leadtimematrixheading">
                    <h2>Project Lead Time Matrix</h2>
                    <p >
                      <b>Last Updated:{" "}</b>
                      {project?.lastNotificationSentAt && !isNaN(new Date(project?.lastNotificationSentAt))
                        ? formatDate(project?.lastNotificationSentAt)
                        : "Not Updated"}
                    </p>
                    {Array.isArray(matrix) && matrix.length > 0 ? (
                      <div className="leadbuttons">
                        <button
                          className="leadtimematrixheadingbutton"
                          disabled={notifyCustomerLoading}
                          onClick={() => handleToNotifyCustomer()}
                        >
                          {notifyCustomerLoading ? <>Notify customer <SpinnerLoader size={10} /></> : "Notify customer"}
                        </button>

                        <button
                          className="leadtimematrixheadingbutton"
                          onClick={openPreview}
                        >
                          Preview
                        </button>
                      </div>
                    ) : null}
                  </div>
                  <div className="matrixTableMain">
                    <table className="matrix-table">
                      {items.length > 0 ? (
                        <thead>
                          <tr>
                            <th>Manufacturer Name</th>
                            <th>Description</th>
                            <th title="Estimated Departure">ETD</th>
                            <th title="Expected Arrival">ETA</th>
                            <th>Arrival</th>
                            <th>Status</th>
                            <th>Actions</th>
                          </tr>
                        </thead>

                      ) : null}

                      <tbody>
                        {items.map((item, index) => {
                          const isEditable = editableRows[index] || !item.id;

                          const handleSave = () => {

                            if (
                              !item.itemName
                            ) {
                              return toast.error(
                                "Manufacturer Name is required."
                              );
                            }
                            if (!validDateChoice(item.expectedDeliveryDate, item.tbdETD)) {
                              return toast.error("For ETD, select a date or mark TBD.");
                            }
                            if (!validDateChoice(item.expectedArrivalDate, item.tbdETA)) {
                              return toast.error("For ETA, select a date or mark TBD.");
                            }
                            if (!validDateChoice(item.arrivalDate, item.tbdArrival)) {
                              return toast.error("For Arrival, select a date or mark TBD.");
                            }
                            const payload = {
                              ...item,
                              expectedDeliveryDate: item.tbdETD ? null : normalizeDate(item.expectedDeliveryDate),
                              expectedArrivalDate: item.tbdETA ? null : normalizeDate(item.expectedArrivalDate),
                              arrivalDate: item.tbdArrival ? null : normalizeDate(item.arrivalDate),
                            };

                            if (item.id) {
                              updateItem(payload);
                            } else {
                              addNewItemToBackend(payload, index);
                            }


                            setEditableRows((prev) => ({
                              ...prev,
                              [index]: false,
                            }));
                          };

                          return (
                            <tr key={item.id || index}>
                              <td>
                                <input
                                  style={{
                                    height: "30px",
                                    borderRadius: "5px",
                                    border: "1px solid #ccc",
                                  }}
                                  value={item.itemName}
                                  disabled={!isEditable}
                                  onChange={(e) =>
                                    handleItemChange(
                                      index,
                                      "itemName",
                                      e.target.value
                                    )
                                  }
                                />
                              </td>
                              <td>
                                <textarea
                                  type="text"
                                  value={item.quantity}
                                  disabled={!isEditable}
                                  onChange={(e) =>
                                    handleItemChange(
                                      index,
                                      "quantity",
                                      e.target.value
                                    )
                                  }
                                  className="item-description-input"
                                  maxLength={50}
                                  style={{
                                    height: "30px",
                                    borderRadius: "5px",
                                    border: "1px solid #ccc",
                                  }}
                                />
                              </td>
                              {/* <td>
                                <label
                                  style={{
                                    display: "flex",
                                    alignItems: "center",
                                    gap: "5px",
                                  }}
                                >
                                  <input
                                    type="checkbox"
                                    checked={item.tbd || false}
                                    disabled={!isEditable}
                                    onChange={(e) => {
                                      const checked = e.target.checked;
                                      handleItemChange(index, "tbd", checked);
                                      if (checked) {
                                        handleItemChange(
                                          index,
                                          "expectedDeliveryDate",
                                          null
                                        );
                                        handleItemChange(
                                          index,
                                          "expectedArrivalDate",
                                          null
                                        );
                                        handleItemChange(
                                          index,
                                          "arrivalDate",
                                          null
                                        );
                                      }
                                    }}
                                  />
                                  TBD
                                </label>
                              </td> */}

                              {/* ETD */}
                              <td>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  <input
                                    type="date"
                                    style={{ height: "30px", borderRadius: "5px", border: "1px solid #ccc" }}
                                    value={toDateStr(item.expectedDeliveryDate)}
                                    disabled={(item.tbdETD ?? false) || !isEditable}
                                    onChange={(e) =>
                                      handleItemChange(index, "expectedDeliveryDate", e.target.value)
                                    }
                                  />
                                  <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                    <input
                                      type="checkbox"
                                      checked={!!(item.tbdETD ?? false)}
                                      disabled={!isEditable}
                                      onChange={(e) => {
                                        const checked = e.target.checked;
                                        handleItemChange(index, "tbdETD", checked);
                                        if (checked) handleItemChange(index, "expectedDeliveryDate", null);
                                      }}
                                    />
                                    TBD
                                  </label>
                                </div>
                              </td>

                              {/* ETA */}
                              <td>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  <input
                                    type="date"
                                    style={{ height: "30px", borderRadius: "5px", border: "1px solid #ccc" }}
                                    value={toDateStr(item.expectedArrivalDate)}
                                    disabled={(item.tbdETA ?? false) || !isEditable}
                                    onChange={(e) =>
                                      handleItemChange(index, "expectedArrivalDate", e.target.value)
                                    }
                                  />
                                  <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                    <input
                                      type="checkbox"
                                      checked={!!(item.tbdETA ?? false)}
                                      disabled={!isEditable}
                                      onChange={(e) => {
                                        const checked = e.target.checked;
                                        handleItemChange(index, "tbdETA", checked);
                                        if (checked) handleItemChange(index, "expectedArrivalDate", null);
                                      }}
                                    />
                                    TBD
                                  </label>
                                </div>
                              </td>

                              {/* Arrival */}
                              <td>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  <input
                                    type="date"
                                    style={{ height: "30px", borderRadius: "5px", border: "1px solid #ccc" }}
                                    value={toDateStr(item.arrivalDate)}
                                    disabled={(item.tbdArrival ?? false) || !isEditable}
                                    onChange={(e) =>
                                      handleItemChange(index, "arrivalDate", e.target.value)
                                    }
                                  />
                                  <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                    <input
                                      type="checkbox"
                                      checked={!!(item.tbdArrival ?? false)}
                                      disabled={!isEditable}
                                      onChange={(e) => {
                                        const checked = e.target.checked;
                                        handleItemChange(index, "tbdArrival", checked);
                                        if (checked) handleItemChange(index, "arrivalDate", null);
                                      }}
                                    />
                                    TBD
                                  </label>
                                </div>
                              </td>

                              <td>
                                <select
                                  value={item.status}
                                  disabled={!isEditable}
                                  onChange={(e) =>
                                    handleItemChange(
                                      index,
                                      "status",
                                      e.target.value
                                    )
                                  }
                                >
                                  <option value="Pending">Pending</option>
                                  <option value="In Transit">In Transit</option>
                                  <option value="Delivered">Delivered</option>
                                  <option value="Installed">Installed</option>
                                  <option value="Arrived">Arrived</option>
                                </select>
                              </td>
                              <td>
                                {item.id ? (
                                  isEditable ? (
                                    <button
                                      disabled={buttonClicked ? true : false}
                                    >
                                      {" "}
                                      {buttonClicked ? (
                                        "..."
                                      ) : (
                                        <span onClick={handleSave}>Save</span>
                                      )}
                                    </button>
                                  ) : (
                                    <>
                                      <button
                                        onClick={() => toggleEditRow(index)}
                                      >
                                        <MdEdit />
                                      </button>
                                      <button onClick={() => deleteItem(item.id)}>
                                        <MdDelete />
                                      </button>
                                      <div
                                        onClick={() =>
                                          itemMarkItemCommentsAsRead(item?.id)
                                        }
                                      >
                                        <button
                                          onClick={() => openItemComment(item.id)}
                                          title="Comment"
                                        >
                                          <FaCommentAlt />
                                          {commentCountsByManufacturerId[
                                            item.id
                                          ] > 0 && (
                                              <span
                                                style={{
                                                  color: "red",
                                                  fontWeight: "bold",
                                                }}
                                              >
                                                (
                                                {
                                                  commentCountsByManufacturerId[
                                                  item.id
                                                  ]
                                                }
                                                )
                                              </span>
                                            )}
                                        </button>
                                      </div>
                                    </>
                                  )
                                ) : (
                                  <>
                                    <button
                                      disabled={buttonClicked ? true : false}
                                    >
                                      {" "}
                                      {buttonClicked ? (
                                        "Saving"
                                      ) : (
                                        <span onClick={handleSave}>Save</span>
                                      )}
                                    </button>
                                    <button onClick={() => removeRow(index)}>
                                      Remove
                                    </button>
                                  </>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  <div
                    className="button1"
                    style={{ display: "flex", justifyContent: "space-between" }}
                  >
                    <button className="ledbutton" onClick={handleAddItemRow}>
                      + Add Row
                    </button>
                  </div>
                </div>
              )}
              {activeTab === "punchlist" && (
                <div className="project-info-card">

                  <div style={{ textAlign: "right", marginBottom: "1rem" }} className="xyzxc">
                    <p >
                      <b>Last Updated:{" "}</b>
                      {project.lastNotificationSentAt && !isNaN(new Date(project.lastNotificationSentAt))
                        ? formatDate(project.lastNotificationSentAt)
                        : "Not Updated"}
                    </p>
                    <button
                      className="leadtimematrixheadingbutton"
                      disabled={notifyCustomerLoading}
                      onClick={() => handleToNotifyCustomerofPunchList()}
                    >
                      {notifyCustomerLoading ? (
                        <>
                          Notify customer <SpinnerLoader size={10} />
                        </>
                      ) : (
                        "Notify customer"
                      )}
                    </button>
                    <button
                      className="ledbutton"
                      onClick={() => setShowPunchModal(true)}
                    >
                      + Add
                    </button>

                  </div>

                  {showPunchModal && (
                    <div className="modal-overlay1">
                      <div className="modal-content">
                        <h3>Add Punch List Issue</h3>
                        <label>Manufacturer</label>
                        <select
                          value={newIssue.projectItemId}
                          onChange={(e) =>
                            setNewIssue({
                              ...newIssue,
                              projectItemId: e.target.value,
                            })
                          }
                        >
                          <option value="">Select</option>
                          {projectItems.map((item) => (
                            <option key={item.id} value={item.id}>
                              {item.itemName}
                            </option>
                          ))}
                        </select>
                        <label>Title</label>
                        <input
                          value={newIssue.title}
                          onChange={(e) =>
                            setNewIssue({ ...newIssue, title: e.target.value })
                          }
                          maxLength={30}
                        />

                        <label>Description</label>
                        <textarea
                          value={newIssue.issueDescription}
                          onChange={(e) =>
                            setNewIssue({
                              ...newIssue,
                              issueDescription: e.target.value,
                            })
                          }
                          minLength={50}
                          maxLength={200}
                        />

                        <label>Upload Files</label>
                        <input
                          type="file"
                          accept="image/*"
                          multiple
                          onChange={handleImageSelect}
                        />
                        {newIssue.productImages.length > 0 && (
                          <div className="image-preview-grid">
                            {newIssue.productImages.map((img, index) => {
                              const imageUrl =
                                typeof img === "string"
                                  ? img
                                  : URL.createObjectURL(img);

                              return (
                                <div key={index} className="preview-wrapper">
                                  <img
                                    src={imageUrl}
                                    alt={`preview-${index}`}
                                    className="preview-img"
                                  />
                                  <span
                                    className="remove-btn"
                                    onClick={() => removeImage(index)}
                                  >
                                    ×
                                  </span>
                                </div>
                              );
                            })}
                          </div>
                        )}

                        <div className="modal-actions">
                          <button
                            className="submit-btn"
                            onClick={async () => {
                              try {
                                // Validation
                                if (
                                  !newIssue.projectItemId ||
                                  !newIssue.title.trim() ||
                                  !newIssue.issueDescription.trim()
                                ) {
                                  toast.error("All fields are required.");
                                  return;
                                }

                                setPunchListLoading(true);
                                const formData = new FormData();
                                const stored = JSON.parse(
                                  localStorage.getItem("user")
                                );
                                const storedUser = stored?.user;
                                const storedCustomer = JSON.parse(
                                  localStorage.getItem("customer")
                                );

                                let createdById = "";
                                let createdByType = "";

                                if (storedUser?.id) {
                                  createdById = storedUser.id;
                                  createdByType = "user";
                                } else if (storedCustomer?.id) {
                                  createdById = storedCustomer.id;
                                  createdByType = "customer";
                                } else {
                                  toast.error("User not logged in");
                                  setPunchListLoading(false);
                                  return;
                                }

                                formData.append("createdById", createdById);
                                formData.append("createdByType", createdByType);
                                formData.append("title", newIssue.title);
                                formData.append(
                                  "issueDescription",
                                  newIssue.issueDescription
                                );
                                formData.append(
                                  "projectItemId",
                                  newIssue.projectItemId
                                );
                                formData.append("projectId", projectId);
                                formData.append(
                                  "category",
                                  projectItems.find(
                                    (p) => p.id == newIssue.projectItemId
                                  )?.itemName || ""
                                );

                                for (let file of newIssue.productImages) {
                                  formData.append("productImages", file);
                                }

                                await axios.post(
                                  `${url}/projects/${projectId}/punch-list`,
                                  formData
                                );
                                toast.success("Issue added!");
                                setShowPunchModal(false);
                                setNewIssue({
                                  title: "",
                                  issueDescription: "",
                                  projectItemId: "",
                                  productImages: [],
                                });

                                const res = await axios.get(
                                  `${url}/projects/${projectId}/punch-list`
                                );
                                setPunchList(
                                  res.data.map((issue) => ({
                                    ...issue,
                                    productImages:
                                      typeof issue.productImages === "string"
                                        ? JSON.parse(issue.productImages)
                                        : issue.productImages,
                                  }))
                                );
                              } catch (err) {
                                console.error("Failed to add issue", err);
                              } finally {
                                setPunchListLoading(false);
                              }
                            }}
                            disabled={punchListLoading}
                          >
                            {punchListLoading ? <SpinnerLoader /> : "Submit"}
                          </button>

                          <button
                            className="cancel-btn"
                            onClick={() => setShowPunchModal(false)}
                          >
                            Cancel
                          </button>
                        </div>
                      </div>
                    </div>
                  )}

                  <h2>Punch List</h2>

                  {punchList.length === 0 ? (
                    <p>No punch list issues found.</p>
                  ) : (
                    <div className="punch-list-grid compact-style">
                      {punchList.map((issue, idx) => {
                        const files = Array.isArray(issue.productImages)
                          ? issue.productImages
                          : [];
                        const currentStatus = statusMap[issue.id] || "Pending";

                        return (
                          <div
                            className="punch-card-small"
                            key={issue.id || idx}
                          >
                            <div className="punch-card-top">
                              <h4>{issue.title}</h4>
                              <span
                                className={`status-badge ${issue.status?.toLowerCase() || "pending"
                                  }`}
                              >
                                {issue.status || "Pending"}
                              </span>
                            </div>

                            <div className="punch-mini-info">
                              <p>
                                <strong>Manufacturer:</strong>{" "}
                                {issue.item?.itemName || "N/A"}
                              </p>
                              <p>
                                <strong>Description:</strong>{" "}
                                {issue.issueDescription}
                              </p>
                            </div>

                            {/* Status Dropdown + Update Button */}
                            <div className="status-control">
                              <select
                                className="status-dropdown"
                                value={currentStatus}
                                onChange={(e) =>
                                  setStatusMap((prev) => ({
                                    ...prev,
                                    [issue.id]: e.target.value,
                                  }))
                                }
                              >
                                <option value="Pending">Pending</option>
                                <option value="Resolved">Resolved</option>
                                <option value="Rejected">Rejected</option>
                              </select>
                              <button
                                className="update-status-btn"
                                onClick={() =>
                                  handlePunchStatusUpdate(
                                    issue.id,
                                    currentStatus
                                  )
                                }
                              >
                                Save
                              </button>
                            </div>

                            {/* Attachments */}
                            {files.length > 0 && (
                              <div className="punch-attachments">
                                {files.map((file, i) => {
                                  const isPDF = file
                                    .toLowerCase()
                                    .endsWith(".pdf");
                                  const fileUrl = `${url2}/${file}`;
                                  return (
                                    <div
                                      key={i}
                                      className="punch-attachment-preview"
                                    >
                                      {isPDF ? (
                                        <iframe
                                          src={fileUrl}
                                          title={`PDF-${i}`}
                                          className="file-mini-pdf"
                                        />
                                      ) : (
                                        <a
                                          href={fileUrl}
                                          target="_blank"
                                          rel="noopener noreferrer"
                                        >
                                          <img
                                            src={fileUrl}
                                            alt={`Img-${i}`}
                                            className="file-mini-img"
                                          />
                                        </a>
                                      )}
                                    </div>
                                  );
                                })}
                              </div>
                            )}
                            <p>
                              <strong>Created At:</strong>{" "}
                              {new Date(issue.createdAt).toLocaleString()}
                            </p>

                            <p>
                              <strong>Created By:</strong>{" "}
                              {issue.createdByType === "user"
                                ? `${issue.creatorUser?.firstName || ""} ${issue.creatorUser?.lastName || ""
                                }`
                                : issue.creatorCustomer?.full_name || "N/A"}
                            </p>
                            <div
                              onClick={() =>
                                markPunchListItemCommentsAsRead(issue.id)
                              }
                            >
                              <button
                                className="comment-btna"
                                onClick={() => openPunchComment(issue.id)}
                                title="Comment"
                              >
                                <FaCommentAlt />
                                {commentCountsByIssueId[issue.id] > 0 && (
                                  <span
                                    style={{ color: "red", fontWeight: "bold" }}
                                  >
                                    ({commentCountsByIssueId[issue.id]})
                                  </span>
                                )}
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              )}
              {activeTab === "invoice" && (
                <InvoiceManagement projectId={projectId} />
              )}
              {activeTab === "settings" && (
                <div className="project-info-card">
                  <h2>Settings</h2>
                  <div className="info-group">
                    <strong>Client View:</strong>{" "}
                    {project.allowClientView ? "Allowed" : "Disabled"}
                  </div>
                  <div className="info-group">
                    <strong>Comments:</strong>{" "}
                    {project.allowComments ? "Allowed" : "Disabled"}
                  </div>
                  <div className="info-group">
                    <strong>Email Notifications:</strong>{" "}
                    {project.enableNotifications ? "Allowed" : "Disabled"}
                  </div>
                </div>
              )}
            </div>
          </>
        )}
      </div>

      {isPreviewAllOpen && (
        <div className="ltm-overlay">
          <div className="ltm-modal">
            <h3 className="ltm-title">Lead Time Matrix — All Items</h3>

            <div className="ltm-table-wrap">
              <table className="ltm-table">
                <thead>
                  <tr>
                    <th>Manufacturer Name</th>
                    <th>Description</th>
                    <th title="Estimated Departure">ETD</th>
                    <th title="Expected Arrival">ETA</th>
                    <th>Arrival</th>
                    <th>Status</th>
                  </tr>
                </thead>

                <tbody>
                  {(matrix || []).map((i) => (
                    <tr key={i.id || `${i.itemName}-${Math.random()}`}>
                      <td>{i.itemName || ""}</td>
                      <td className="ltm-desc">{i.quantity || ""}</td>
                      <td>{(i.tbdETD ? "TBD" : toDateStr(i.expectedDeliveryDate))}</td>
                      <td>{(i.tbdETA ? "TBD" : toDateStr(i.expectedArrivalDate))}</td>
                      <td>{(i.tbdArrival ? "TBD" : toDateStr(i.arrivalDate))}</td>

                      <td>{i.status || ""}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="ltm-actions">
              <button className="ltm-btn" onClick={handleDownloadLeadTimeExcel}>
                Download Excel
              </button>
              <button
                className="ltm-btn"
                onClick={handleDownloadLeadTimePDF}
                disabled={!items?.length}
              >
                Download PDF
              </button>
              <button
                className="ltm-btn ltm-btn--secondary"
                onClick={() => setIsPreviewAllOpen(false)}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}



      {/*team comment canvas*/}
      <Offcanvas
        isOpen={isOffcanvasOpen}
        closeOffcanvas={closeOffcanvas}
        getLatestComment={fetchComments}
      >
        <div className="right-panel">
          <div
            className="comments-list"
            style={{
              maxHeight: "500px",
              display: "flex",
              flexDirection: "column",
            }}
          >
            {/* Grouped Comments by Date */}
            {Object.keys(groupedComments)
              .sort((a, b) => new Date(a) - new Date(b))
              .map((date) => (
                <div key={date} className="comment-date-group">
                  <p className="comment-date">{date}</p>
                  {groupedComments[date]
                    .sort(
                      (a, b) => new Date(a.createdAt) - new Date(b.createdAt)
                    ) // Oldest to newest in each group
                    .map((comment) => (
                      <div key={comment.id}>
                        <div className="whatsapp-comment-box">
                          <div className="whatsapp-comment-user-info">
                            <img
                              src={
                                comment?.profileImage
                                  ? `${url2}/${comment.profileImage}`
                                  : `${process.env.PUBLIC_URL}/assets/Default_pfp.jpg`
                              }
                              alt="User"
                              className="whatsapp-comment-user-avatar"
                            />
                            <div>
                              <p className="whatsapp-comment-author">
                                {comment?.name} ({comment?.userRole})
                              </p>
                            </div>
                          </div>
                          <p className="whatsapp-comment-text">
                            {comment.comment}
                          </p>
                          <p className="whatsapp-comment-meta">
                            {new Date(comment.createdAt).toLocaleTimeString()}
                          </p>
                        </div>
                      </div>
                    ))}
                  <div ref={scrollRef} />
                </div>
              ))}
          </div>
        </div>

        <div className="whatsapp-comment-form">
          <textarea
            value={newCommentText}
            onChange={(e) => setNewCommentText(e.target.value)}
            className="whatsapp-comment-input"
            placeholder="Write your comment..."
          />
          <button onClick={handleAddComment} className="whatsapp-submit-btn">
            {<FaTelegramPlane />}
          </button>
        </div>
      </Offcanvas>
      {/*item comment canvas*/}
      <Offcanvas
        isOpen={isItemCanvasOpen}
        closeOffcanvas={() => setIsItemCanvasOpen(false)}
        getLatestComment={() => openItemComment(selectedItemId)}
      >
        <div className="right-panel">
          <div
            className="comments-list"
            style={{
              maxHeight: "500px",
              display: "flex",
              flexDirection: "column",
            }}
          >
            {Object.keys(groupedItemComments)
              .sort((a, b) => new Date(a) - new Date(b))
              .map((date) => (
                <div key={date} className="comment-date-group">
                  <p className="comment-date">{date}</p>
                  {groupedItemComments[date]
                    .sort(
                      (a, b) => new Date(a.createdAt) - new Date(b.createdAt)
                    )
                    .map((comment) => (
                      <div key={comment.id} className="whatsapp-comment-box">
                        <div className="whatsapp-comment-user-info">
                          <img
                            src={
                              comment.profileImage
                                ? `${url2}/${comment.profileImage}`
                                : `${process.env.PUBLIC_URL}/assets/Default_pfp.jpg`
                            }
                            alt="User"
                            className="whatsapp-comment-user-avatar"
                          />
                          <div>
                            <p className="whatsapp-comment-author">
                              {comment?.createdByName && comment?.userRole
                                ? `${comment.createdByName} (${comment.userRole})`
                                : "Customer"}
                            </p>
                          </div>
                        </div>
                        <p className="whatsapp-comment-text">
                          {comment.comment}
                        </p>
                        <p className="whatsapp-comment-meta">
                          {new Date(comment.createdAt).toLocaleTimeString()}
                        </p>
                      </div>
                    ))}
                  <div ref={scrollRef} />
                </div>
              ))}
          </div>
        </div>

        <div className="whatsapp-comment-form">
          <textarea
            value={itemCommentText}
            onChange={(e) => setItemCommentText(e.target.value)}
            className="whatsapp-comment-input"
            placeholder="Write your comment..."
          />
          <button
            onClick={handleAddItemComment}
            className="whatsapp-submit-btn"
          >
            {<FaTelegramPlane />}
          </button>
        </div>
      </Offcanvas>
      {/*punch comment canvas*/}
      <Offcanvas
        isOpen={isPunchCanvasOpen}
        closeOffcanvas={() => setIsPunchCanvasOpen(false)}
        getLatestComment={() => openPunchComment(selectedPunchItemId)}
      >
        <div className="right-panel">
          <div
            className="comments-list"
            style={{
              maxHeight: "500px",
              display: "flex",
              flexDirection: "column",
            }}
          >
            {Object.keys(groupedPunchComments)
              .sort((a, b) => new Date(a) - new Date(b)) // Oldest date group first
              .map((date) => (
                <div key={date} className="comment-date-group">
                  <p className="comment-date">{date}</p>

                  {groupedPunchComments[date]
                    .sort(
                      (a, b) => new Date(a.createdAt) - new Date(b.createdAt)
                    ) // Oldest to newest inside each group
                    .map((comment) => {
                      const isUser = !!comment.user;
                      const creator = isUser ? comment.user : comment.customer;

                      return (
                        <div key={comment.id} className="whatsapp-comment-box">
                          <div className="whatsapp-comment-user-info">
                            <img
                              src={
                                comment.profileImage
                                  ? `${url2}/${comment.profileImage}`
                                  : `${process.env.PUBLIC_URL}/assets/Default_pfp.jpg`
                              }
                              alt="User"
                              className="whatsapp-comment-user-avatar"
                            />
                            <div>
                              <p className="whatsapp-comment-author">
                                {comment.createdByType === 'customer'
                                  ? `Customer${comment?.name ? ' · ' + comment.name : ''}`
                                  : `${comment.name || ''}${comment.userRole ? ` (${comment.userRole})` : ''}`
                                }
                              </p>
                            </div>
                          </div>
                          <p className="whatsapp-comment-text">
                            {comment.comment}
                          </p>
                          <p className="whatsapp-comment-meta">
                            {new Date(comment.createdAt).toLocaleTimeString()}
                          </p>
                        </div>
                      );
                    })}

                  {/* Scroll to latest (last) comment */}
                  <div ref={scrollRef} />
                </div>
              ))}
          </div>
        </div>

        <div className="whatsapp-comment-form">
          <textarea
            value={punchCommentText}
            onChange={(e) => setPunchCommentText(e.target.value)}
            className="whatsapp-comment-input"
            placeholder="Write your comment..."
          />
          <button
            onClick={handleAddPunchComment}
            className="whatsapp-submit-btn"
          >
            {<FaTelegramPlane />}
          </button>
        </div>
      </Offcanvas>
    </Layout>
  );
};

export default ProjectDetails;
