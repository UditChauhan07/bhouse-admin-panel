import React, { useState, useEffect } from "react";
import Layout from "../components/Layout";
import { getRoles, deleteRole } from "../lib/api";import { FaEdit, FaTrash } from "react-icons/fa";

import { useNavigate } from "react-router-dom";
import Swal from "sweetalert2";
import "../styles/Roles.css";
import useRolePermissions from "../hooks/useRolePermissions";
import Loader from "../components/Loader";

const rolePermissionLevels = {
  "Admin" : 1 , 
  "Super Admin": 1,
  "Account Manager": 2,
  "Sr. Designer": 3,
  "Operations": 4,
  "Junior Designer" : 5 ,
  "Lead Installer": 6,
};

const Roles = () => {
  const [roles, setRoles] = useState([]);
  const [search, setSearch] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [loading , setLoading] = useState(true)
  const itemsPerPage = 8;
  const navigate = useNavigate();

  const user = JSON.parse(localStorage.getItem("user"));
  const loggedInUserRole = (user?.user?.userRole || "").trim();
  const loggedInUserId = user?.user?.id;
  const roleId = user?.user?.roleId;
  const { rolePermissions } = useRolePermissions(roleId);


  useEffect(() => {
    fetchRoles();
  }, []);
  const fetchRoles = async () => {
    try {
      const data = await getRoles();
      const allRoles = data?.data || [];
       if(data){
        setLoading(false)
       }
      const loggedUserId = parseInt(loggedInUserId, 10);
      const userLevel = rolePermissionLevels[loggedInUserRole] || 6; // Fetch user level

      const filteredRoles = allRoles.filter((role) => {
        const roleLevel = role.defaultPermissionLevel || 7;
        const createdById = parseInt(role.createdBy, 10); // Ensure it's a number

        if (userLevel === 1) return true; // Super Admin can see all

        if (userLevel === 6) {
          const isOwnRole = roleLevel === 6 && createdById === loggedUserId;
          return isOwnRole;
        }

        // ✅ Correct condition to show only lower level roles or roles created by the user
        const shouldShow = (roleLevel > userLevel && roleLevel < 7) || createdById === loggedUserId;
        return shouldShow;
      });



      setRoles(filteredRoles);
    } catch (error) {
      console.error("Error fetching roles:", error);
      setRoles([]);
    }
  };

  const handleDelete = async (id) => {
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#d33",
      cancelButtonColor: "#3085d6",
      confirmButtonText: "Yes, delete it!",
    }).then(async (result) => {
      if (result.isConfirmed) {
        await deleteRole(id);
        setRoles(roles.filter((role) => role.id !== id));
        Swal.fire({
          icon: "success",
          title: "Deleted!",
          text: "The role has been deleted.",
          timer: 600,
          showConfirmButton: false
        });
      }
    });
  };

  // ✅ Apply search filter
  const filteredRoles = (roles || []).filter((role) =>
    role?.title?.toLowerCase().includes(search.toLowerCase())
  );

  // ✅ Pagination Logic
  const totalPages = Math.ceil(filteredRoles.length / itemsPerPage);
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = filteredRoles.slice(indexOfFirstItem, indexOfLastItem);

  return (
    <Layout>
      <div className="roles-container">
        <h2 className="table-header">Roles</h2>
        <div className="roles-header">
          {rolePermissions?.Roles?.create && (
            <button className="add-role-btn" onClick={() => navigate("/create-role")}>
              Add Role
            </button>
          )}

          <input
            type="text"
            placeholder="Search roles..."
            value={search}
            onChange={(e) => {
              setSearch(e.target.value);
              setCurrentPage(1);
            }}
            className="search-input"
          />
        </div>
        {loading ? <Loader/> :
        <table className="roles-table">
          <thead>
            <tr>
              <th>Sr. No</th>
              <th>Title</th>
              <th>Description</th>
              {(rolePermissions?.Roles?.edit || rolePermissions?.Roles?.delete) && <th>Actions</th>}

            </tr>
          </thead>
          
           <tbody>
           {currentItems.length > 0 ? (
             currentItems.map((role, index) => (
               <tr key={index}>
                <td>{(currentPage - 1) * itemsPerPage + index + 1}</td> 
                 <td>{role.title}</td>
                 <td>{role.description || "N/A"}</td>
                 {(rolePermissions?.Roles?.edit || rolePermissions?.Roles?.delete) && (
                   <td className="actions1">
                  {rolePermissions?.Roles?.edit && role.title !== "Super Admins " ? (
  <FaEdit
    style={{
      color: "#004680",
      fontSize: "24px",
    }}
    className="edit-icon"
    title="Edit"
    onClick={() => navigate(`/edit-role/${role.id}`)}
  />
) : "NA"}
                     {/* {rolePermissions?.Roles?.delete && (
                       <FaTrash
                       style={{
                        color: "#004680" ,
                         fontSize : "20px"
                     }}
                         title="Delete"
                         onClick={() => handleDelete(role.id)}
                       />
                     )} */}
                   </td>
                 )}

               </tr>
             ))
           ) : (
             <tr>
               <td colSpan="4" className="no-data">No Data Found</td>
             </tr>
           )}
         </tbody>
        
         
        </table>
      }

        {/* Pagination */}
        {totalPages > 1 && (
          <div className="pagination">
            <button
              className="icon-btn"
              onClick={() => setCurrentPage(currentPage - 1)}
              disabled={currentPage === 1}
            >
              ◀
            </button>

            {[...Array(totalPages)].map((_, index) => (
              <button
                key={index}
                className={`page-btn ${currentPage === index + 1 ? "active" : ""}`}
                onClick={() => setCurrentPage(index + 1)}
              >
                {index + 1}
              </button>
            ))}

            <button
              className="icon-btn"
              onClick={() => setCurrentPage(currentPage + 1)}
              disabled={currentPage === totalPages}
            >
              ▶
            </button>
          </div>
        )}
      </div>
    </Layout>
  );
};

export default Roles;
